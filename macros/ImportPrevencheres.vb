Option Explicit

' ========= IMPORT EXCEL -> MS PROJECT (VISION OMX) =========
' Colonnes Excel (A:L)
' A Nom
' B Quantités
' C Nombre de personnes
' D Heures Budgetées
' E Zone
' F Sous-Zone
' G Tranche
' H Lot
' I Entreprise
' J Qualité   (CQ / TACHE / vide)
' K Niveau    (SZ / OND / vide pour titres)
' L Onduleur  (obligatoire si Niveau=OND)

Sub Import_Taches_Simples_AvecTitre()

    Dim xlApp As Object, xlBook As Object, xlSheet As Object
    Dim pjApp As MSProject.Application, pjProj As MSProject.Project
    Dim i As Long, lastRow As Long
    Dim t As Task, tCQ As Task, tRoot As Task, tGroup As Task
    Dim a As Assignment
    Dim fichierExcel As String

    ' ==== SELECTION DU FICHIER VIA SELECTEUR NATIF ====
    Dim xlTempApp As Object
    Set xlTempApp = CreateObject("Excel.Application")
    xlTempApp.Visible = False

    With xlTempApp.FileDialog(msoFileDialogFilePicker)
        .Title = "Sélectionnez le fichier Excel à importer"
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
    ' Mapping tags (modifiable)
    ' Text1 = Tranche
    ' Text2 = Zone
    ' Text3 = Sous-Zone
    ' Text4 = Lot
    ' Text5 = Entreprise
    ' Text6 = Niveau
    ' Text7 = Onduleur
    pjApp.CustomFieldRename pjCustomTaskText1, "Tranche"
    pjApp.CustomFieldRename pjCustomTaskText2, "Zone"
    pjApp.CustomFieldRename pjCustomTaskText3, "Sous-Zone"
    pjApp.CustomFieldRename pjCustomTaskText4, "Lot"
    pjApp.CustomFieldRename pjCustomTaskText5, "Entreprise"
    pjApp.CustomFieldRename pjCustomTaskText6, "Niveau"
    pjApp.CustomFieldRename pjCustomTaskText7, "Onduleur"

    ' ==== AJOUT DU TITRE DE PROJET (CELLULE A2) ====
    Set tRoot = pjProj.Tasks.Add(Name:=CStr(xlSheet.Cells(2, 1).Value), Before:=1)
    tRoot.Manual = False
    tRoot.OutlineLevel = 1

    ' Le groupe courant (titre de section) : au début on met Root
    Set tGroup = tRoot

    ' ==== CONFIGURATION PROJET ====
    pjProj.DefaultTaskType = pjFixedWork
    pjProj.ScheduleFromStart = True
    pjProj.DefaultEffortDriven = True

    ' ==== CALENDRIER STANDARD (optionnel, garde ta version) ====
    On Error Resume Next
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
    On Error GoTo 0

    ' ==== RESSOURCES STANDARD ====
    Dim rProd As Resource
    Set rProd = GetOrCreateWorkResource("Monteurs") ' ressource production générique

    Dim rCQWork As Resource
    Set rCQWork = GetOrCreateWorkResource("Controleur CQ") ' CQ OMX (Work)

    Dim rCQMat As Resource
    Set rCQMat = GetOrCreateMaterialResource("CQ") ' compteur CQ si besoin

    lastRow = xlSheet.Cells(xlSheet.Rows.Count, 1).End(-4162).Row

    ' ==== BOUCLE LIGNES ====
    For i = 3 To lastRow

        Dim nom As String, qte As Variant, pers As Variant, h As Variant
        Dim zone As String, sousZone As String, tranche As String, lot As String, entreprise As String
        Dim qualite As String, niveau As String, onduleur As String

        nom = Trim$(CStr(xlSheet.Cells(i, 1).Value))
        qte = xlSheet.Cells(i, 2).Value
        pers = xlSheet.Cells(i, 3).Value
        h = xlSheet.Cells(i, 4).Value

        zone = Trim$(CStr(xlSheet.Cells(i, 5).Value))         ' E
        sousZone = Trim$(CStr(xlSheet.Cells(i, 6).Value))     ' F
        tranche = Trim$(CStr(xlSheet.Cells(i, 7).Value))      ' G
        lot = Trim$(CStr(xlSheet.Cells(i, 8).Value))          ' H
        entreprise = Trim$(CStr(xlSheet.Cells(i, 9).Value))   ' I

        qualite = UCase$(Trim$(CStr(xlSheet.Cells(i, 10).Value))) ' J
        niveau = UCase$(Trim$(CStr(xlSheet.Cells(i, 11).Value)))  ' K
        onduleur = UCase$(Trim$(CStr(xlSheet.Cells(i, 12).Value))) ' L

        If nom = "" Then
            GoTo NextRow
        End If

        ' Détection TITRE: si pas de data (heures, qté, pers) et pas de tags structurels
        ' => on crée une tâche récap sous Root, et on attache les tâches suivantes dessous
        If IsEmptyOrZero(qte) And IsEmptyOrZero(pers) And IsEmptyOrZero(h) _
           And zone = "" And sousZone = "" And tranche = "" And lot = "" And entreprise = "" _
           And qualite = "" And niveau = "" And onduleur = "" Then

            Set tGroup = pjProj.Tasks.Add(nom)
            tGroup.Manual = False
            tGroup.OutlineLevel = 2
            GoTo NextRow
        End If

        ' Validation Niveau/Onduleur
        If niveau = "OND" And onduleur = "" Then
            MsgBox "Ligne " & i & " : Niveau=OND mais Onduleur vide pour la tâche '" & nom & "'.", vbExclamation
        End If

        ' === Création tâche principale ===
        Set t = pjProj.Tasks.Add(nom)
        t.Manual = False
        t.Calendar = ActiveProject.BaseCalendars("Standard")
        t.OutlineLevel = 3 ' sous le groupe (niveau 2)

        ' Tags -> champs texte
        t.Text1 = tranche
        t.Text2 = zone
        t.Text3 = sousZone
        t.Text4 = lot
        t.Text5 = entreprise
        t.Text6 = niveau
        t.Text7 = onduleur

        ' === Affectation production (heures/personnes) ===
        If IsNumeric(h) And CDbl(h) > 0 Then
            Dim nbPers As Double
            nbPers = IIf(IsNumeric(pers) And CDbl(pers) > 0, CDbl(pers), 1#)

            Set a = t.Assignments.Add(ResourceID:=rProd.ID)
            a.Units = nbPers
            a.Work = CDbl(h) * 60 ' minutes
        End If

        ' === Quantité (consommable) ===
        If IsNumeric(qte) And CDbl(qte) > 0 Then
            Dim rMat As Resource
            Set rMat = Nothing
            ' Option 1 : une ressource matériau par NOM DE TÂCHE (comme avant) -> pas idéal mais V0
            Set rMat = GetOrCreateMaterialResource(nom)
            Set a = t.Assignments.Add(ResourceID:=rMat.ID)
            
            ' ✅ CORRECTIF : répartir la quantité sur la durée de la tâche
            ' Pour les ressources Material, Work représente la durée de consommation
            If t.Duration > 0 Then
                a.Work = t.Duration ' Durée de la tâche en minutes
                ' Units = taux de consommation par minute
                a.Units = CDbl(qte) / (t.Duration / 60) ' Quantité par heure
            Else
                ' Si pas de durée, consommation ponctuelle
                a.Units = CDbl(qte)
            End If
        End If

        ' === LOGIQUE CQ HYBRIDE ===
        ' Règle souhaitée :
        ' - si tâche réalisée par SST/Externe -> tâche CQ séparée
        ' - si tâche réalisée par OMX -> CQ en ressource (Work recommandé)
        Dim isOmx As Boolean
        isOmx = (UCase$(entreprise) = "OMX" Or UCase$(entreprise) = "OMEXOM")

        If qualite = "CQ" Then
            If isOmx Then
                ' CQ intégré : ressource Travail "Controleur CQ" (cadence)
                Set a = t.Assignments.Add(ResourceID:=rCQWork.ID)
                a.Units = 1 ' V0
                ' Si tu veux que CQ consomme des heures, il faut une règle.
                ' Ici, si la colonne Heures est renseignée, on la met en work CQ et on laisse "Monteurs" vide.
                ' Sinon, on laisse à 0 (simple marquage de besoin).
            Else
                ' CQ SST : malgré "CQ", on force une tâche CQ séparée (vision Simon)
                CreateCQTask pjProj, t, tCQ, rCQWork, tranche, zone, sousZone, lot, entreprise, niveau, onduleur
            End If

        ElseIf qualite = "TACHE" Or qualite = "TÂCHE" Then
            ' CQ explicitement en tâche
            CreateCQTask pjProj, t, tCQ, rCQWork, tranche, zone, sousZone, lot, entreprise, niveau, onduleur
        End If

NextRow:
    Next i

    xlBook.Close SaveChanges:=False
    xlApp.Quit
    Set xlApp = Nothing

    MsgBox "Import terminé. Vérifie les tags (Text1..Text7) et la logique CQ (OMEXOM vs SST).", vbInformation

End Sub

' ===== Helper: création tâche CQ + lien FS + tags =====
Private Sub CreateCQTask(ByVal pjProj As MSProject.Project, ByVal tMain As Task, ByRef tCQ As Task, _
                         ByVal rCQWork As Resource, ByVal tranche As String, ByVal zone As String, _
                         ByVal sousZone As String, ByVal lot As String, ByVal entreprise As String, _
                         ByVal niveau As String, ByVal onduleur As String)

    Dim a As Assignment

    Set tCQ = pjProj.Tasks.Add("Contrôle Qualité - " & tMain.Name)
    tCQ.Manual = False
    tCQ.Calendar = ActiveProject.BaseCalendars("Standard")
    tCQ.OutlineLevel = tMain.OutlineLevel ' même niveau visuel

    ' Tags CQ
    tCQ.Text1 = tranche
    tCQ.Text2 = zone
    tCQ.Text3 = sousZone
    tCQ.Text4 = lot
    tCQ.Text5 = "OMEXOM" ' CQ porté par OMX (contrôleur)
    tCQ.Text6 = niveau
    tCQ.Text7 = onduleur

    ' Ressource CQ (Work)
    Set a = tCQ.Assignments.Add(ResourceID:=rCQWork.ID)
    a.Units = 1

End Sub

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


