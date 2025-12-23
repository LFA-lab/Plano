Sub Import_Taches_Simples_AvecTitre()

    Dim xlApp As Object, xlBook As Object, xlSheet As Object
    Dim pjApp As MSProject.Application, pjProj As MSProject.Project
    Dim i As Long, lastRow As Long
    Dim t As Task, a As Assignment
    Dim fichierExcel As String

    ' ==== S�LECTION DU FICHIER VIA INSTANCE EXCEL TEMPORAIRE ====
    Dim xlTempApp As Object
    Set xlTempApp = CreateObject("Excel.Application")

    With xlTempApp.FileDialog(msoFileDialogFilePicker)
        .Title = "S�lectionnez le fichier Excel � importer"
        .Filters.Clear
        .Filters.Add "Fichiers Excel", "*.xlsx; *.xls"
        .AllowMultiSelect = False
        If .Show = -1 Then
            fichierExcel = .SelectedItems(1)
        Else
            MsgBox "Aucun fichier s�lectionn�. L'importation est annul�e.", vbExclamation
            xlTempApp.Quit
            Set xlTempApp = Nothing
            Exit Sub
        End If
    End With

    xlTempApp.Quit
    Set xlTempApp = Nothing

    ' ==== OUVERTURE D'EXCEL ====
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = False
    Set xlBook = xlApp.Workbooks.Open(fichierExcel)
    Set xlSheet = xlBook.Sheets(1)

    ' ==== OUVERTURE DE MS PROJECT ====
    Set pjApp = MSProject.Application
    pjApp.Visible = True
    pjApp.FileNew
    Set pjProj = pjApp.ActiveProject

    ' ==== AJOUT DU TITRE DE PROJET (CELLULE A2) ====
    Dim tRoot As Task
    Set tRoot = pjProj.Tasks.Add(name:=xlSheet.Cells(2, 1).Value, Before:=1)
    tRoot.Manual = False
    tRoot.Calendar = ActiveProject.BaseCalendars("Standard")
    
    ' ==== VARIABLE DE SUIVI POUR LA TÂCHE RÉCAPITULATIVE COURANTE ====
    Dim tCurrentSummary As Task
    Set tCurrentSummary = Nothing

    ' ==== CONFIGURATION PROJET ====
    pjProj.DefaultTaskType = pjFixedWork
    pjProj.ScheduleFromStart = True
    pjProj.DefaultEffortDriven = True

    ' ==== MODIFICATION DU CALENDRIER "Standard" ====
    With ActiveProject.BaseCalendars("Standard").WorkWeeks
        .Add Start:="01/01/2025", Finish:="01/01/2027", name:="Calendrier Standard"
        With .Item(1)
            Dim j As Integer
            For j = 2 To 6 ' Lundi � vendredi
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

    ' ==== CR�ATION RESSOURCE "Monteurs" ====
    Dim rMonteurs As Resource
    Set rMonteurs = GetOrCreateWorkResource("Monteurs")

    lastRow = xlSheet.Cells(xlSheet.Rows.Count, 1).End(-4162).Row ' fin de la colonne A

    ' ==== BOUCLE T�CHES ====
    For i = 3 To lastRow
        Dim nom As String, qte As Variant, pers As Variant, h As Variant
        Dim strQte As String, strPers As String, strH As String
        nom = xlSheet.Cells(i, 1).Value
        qte = xlSheet.Cells(i, 2).Value
        pers = xlSheet.Cells(i, 3).Value
        h = xlSheet.Cells(i, 4).Value
        
        ' Conversion en string pour test de vide (gère les cas où la cellule contient 0 vs vide)
        strQte = Trim(CStr(qte))
        strPers = Trim(CStr(pers))
        strH = Trim(CStr(h))
        Dim strNom As String
        strNom = Trim(CStr(nom))
        
        ' ==== ÉTAPE 1: TEST LIGNE VIDE ====
        ' Si A, B, C, D sont vides → remise à zéro de la récap courante
        If strNom = "" And strQte = "" And strPers = "" And strH = "" Then
            Set tCurrentSummary = Nothing
            GoTo ContinueLoop
        End If
        
        ' ==== ÉTAPE 2: TEST TÂCHE RÉCAPITULATIVE ====
        ' Colonne A non vide ET colonnes B, C, D vides
        If strNom <> "" And strQte = "" And strPers = "" And strH = "" Then
            ' Créer la tâche récapitulative
            Set t = pjProj.Tasks.Add(nom)
            t.Manual = False
            t.Calendar = ActiveProject.BaseCalendars("Standard")
            
            ' Rattacher à tRoot (toutes les récaps sont au même niveau, enfants de tRoot)
            t.OutlineParent = tRoot
            
            ' Ne PAS assigner de ressource à une tâche récapitulative
            ' Stocker dans la variable de suivi
            Set tCurrentSummary = t
            GoTo ContinueLoop
        End If
        
        ' ==== ÉTAPE 3: TÂCHE DÉTAILLÉE ====
        ' Colonne A non vide ET au moins une de B, C ou D non vide
        If strNom <> "" And (strQte <> "" Or strPers <> "" Or strH <> "") Then
            Set t = pjProj.Tasks.Add(nom)
            t.Manual = False
            t.Calendar = ActiveProject.BaseCalendars("Standard")
            
            ' Si une récap courante existe, la tâche détaillée devient son enfant
            If Not tCurrentSummary Is Nothing Then
                t.OutlineParent = tCurrentSummary
            End If
            
            ' Assigner les ressources comme dans le code actuel
            If IsNumeric(h) And h > 0 Then
                Dim nbPers As Double: nbPers = IIf(IsNumeric(pers) And pers > 0, pers, 1)
                Set a = t.Assignments.Add(ResourceID:=rMonteurs.ID)
                a.Units = nbPers
                a.Work = h * 60
            End If

            If IsNumeric(qte) And qte > 0 Then
                Dim rMat As Resource
                Set rMat = GetOrCreateMaterialResource(nom)
                Set a = t.Assignments.Add(ResourceID:=rMat.ID)
                a.Units = qte
            End If
        End If
        
ContinueLoop:
    Next i

    ' ==== FERMETURE ====
    xlBook.Close False
    xlApp.Quit
    Set xlApp = Nothing
    MsgBox "? Import termin� avec titre de projet, ressources, et calendrier.", vbInformation

End Sub


Function GetOrCreateWorkResource(nom As String) As Resource
    Dim r As Resource
    On Error Resume Next
    Set r = ActiveProject.Resources(nom)
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
    If r Is Nothing Then
        Set r = ActiveProject.Resources.Add(nom)
        r.Type = pjResourceTypeMaterial
    End If
    Set GetOrCreateMaterialResource = r
End Function

