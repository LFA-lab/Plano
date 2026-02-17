Attribute VB_Name = "Module1"
Option Explicit

Sub Import_Taches_Simples_AvecTitre()

    ' ==============================
    ' DECLARATIONS
    ' ==============================
    
    Dim xlApp As Object
    Dim xlBook As Object
    Dim xlSheet As Object
    
    Dim pjApp As MSProject.Application
    Dim pjProj As MSProject.Project
    
    Dim dataArr As Variant
    Dim resourceCache As Object
    
    Dim i As Long, lastRow As Long
    Dim t As Task
    Dim tCQ As Task
    Dim a As Assignment
    
    Dim fichierExcel As String
    Dim oldScreenUpdating As Boolean
    
    ' ==============================
    ' USE CURRENT PROJECT (NO CREATION)
    ' ==============================
    
    Set pjApp = Application
    Set pjProj = pjApp.ActiveProject
    
    If pjProj Is Nothing Then
        MsgBox "Aucun projet actif.", vbCritical
        Exit Sub
    End If
    
    pjApp.DisplayAlerts = False
    
    On Error Resume Next
    oldScreenUpdating = pjApp.ScreenUpdating
    pjApp.ScreenUpdating = False
    On Error GoTo 0
    
    ' ==============================
    ' SELECT EXCEL FILE
    ' ==============================
    
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = False
    
    fichierExcel = xlApp.GetOpenFilename("Fichiers Excel (*.xlsx), *.xlsx")
    
    If fichierExcel = "False" Then
        MsgBox "Import annulé.", vbInformation
        GoTo CleanExit
    End If
    
    Set xlBook = xlApp.Workbooks.Open(fichierExcel)
    Set xlSheet = xlBook.Sheets(1)
    
    ' ==============================
    ' LOAD DATA INTO MEMORY ARRAY
    ' ==============================
    
    lastRow = xlSheet.Cells(xlSheet.Rows.Count, 1).End(-4162).Row ' xlUp
    
    If lastRow < 2 Then
        MsgBox "Fichier Excel vide.", vbExclamation
        GoTo CleanExit
    End If
    
    dataArr = xlSheet.Range("A2:O" & lastRow).Value
    
    ' ==============================
    ' RESOURCE CACHE (FAST LOOKUP)
    ' ==============================
    
    Set resourceCache = CreateObject("Scripting.Dictionary")
    
    Dim r As Resource
    For Each r In pjProj.Resources
        resourceCache(r.Name) = r
    Next r
    
    ' ==============================
    ' RENAME CUSTOM FIELDS
    ' ==============================
    
    pjApp.CustomFieldRename pjCustomTaskText1, "Tranche"
    pjApp.CustomFieldRename pjCustomTaskText2, "Zone"
    pjApp.CustomFieldRename pjCustomTaskText3, "Sous-Zone"
    pjApp.CustomFieldRename pjCustomTaskText4, "Metier"
    pjApp.CustomFieldRename pjCustomTaskText5, "Entreprise"
    pjApp.CustomFieldRename pjCustomTaskText6, "Niveau"
    pjApp.CustomFieldRename pjCustomTaskText7, "Onduleur"
    pjApp.CustomFieldRename pjCustomTaskText8, "PTR"
    
    ' ==============================
    ' CREATE TASKS (IMPORT_OPTIMISE style)
    ' ==============================

    ' Ensure Monteurs resource exists in cache (will be created on demand)
    Dim rMonteurs As Resource
    ' loop through data rows (dataArr row 1 = Excel row 2)
    Dim tGroup As Task
    Set tGroup = Nothing

    For i = 1 To UBound(dataArr, 1)

        Dim nom As String, qte As Variant, pers As Variant, h As Variant
        Dim zone As String, sousZone As String, tranche As String, typ As String, entreprise As String
        Dim qualite As String, niveau As String, onduleur As String, ptr As String
        Dim dateDebutMonteurs As Date, dateFinMonteurs As Date
        Dim hasMonteursAssignment As Boolean

        nom = Trim$(CStr(dataArr(i, 1)))           ' A
        qte = dataArr(i, 2)                        ' B
        pers = dataArr(i, 3)                       ' C
        h = dataArr(i, 4)                          ' D
        zone = Trim$(CStr(dataArr(i, 5)))          ' E
        sousZone = Trim$(CStr(dataArr(i, 6)))      ' F
        tranche = Trim$(CStr(dataArr(i, 7)))       ' G
        typ = Trim$(CStr(dataArr(i, 8)))           ' H
        entreprise = Trim$(CStr(dataArr(i, 9)))    ' I
        qualite = UCase$(Trim$(CStr(dataArr(i, 10)))) ' J
        niveau = UCase$(Trim$(CStr(dataArr(i, 11))))  ' K
        onduleur = UCase$(Trim$(CStr(dataArr(i, 12)))) ' L
        On Error Resume Next
        ptr = Trim$(CStr(dataArr(i, 13)))          ' M (optionnel)
        If Err.Number <> 0 Then ptr = ""
        On Error GoTo 0

        hasMonteursAssignment = False

        If nom = "" Then GoTo NextRow

        ' DETECT TITLE (no qte and no h)
If IsEmptyOrZero(qte) And IsEmptyOrZero(h) Then
    
    Set tGroup = pjProj.Tasks.Add(nom)
    tGroup.Manual = False
    
    ' Indent based on previous task
    If pjProj.Tasks.Count > 1 Then
        On Error Resume Next
        
        If InStr(1, nom, "ZONE", vbTextCompare) > 0 Then
            tGroup.OutlineIndent   ' Level 2
        Else
            tGroup.OutlineIndent   ' Level 2
            tGroup.OutlineIndent   ' Level 3
        End If
        
        On Error GoTo 0
    End If
    
    GoTo NextRow
End If

        ' Create task
        Set t = pjProj.Tasks.Add(nom)
        If t Is Nothing Then GoTo NextRow

        t.Manual = False
        t.Calendar = pjProj.BaseCalendars("Standard")
        t.LevelingCanSplit = False

        ' Tags
        t.Text1 = tranche
        t.Text2 = zone
        t.Text3 = sousZone
        t.Text4 = typ
        t.Text5 = entreprise
        t.Text6 = niveau
        t.Text7 = onduleur
        t.Text8 = ptr

        ' WORK before Units: set task work if hours specified
        Dim workMinutes As Long
        If IsNumeric(h) And CDbl(h) > 0 Then
            workMinutes = CLng(CDbl(h) * 60)
            t.Type = pjFixedWork
            t.Work = workMinutes
        Else
            ' default = 1 day (assume 8h = 480 minutes)
            workMinutes = 480
            t.Type = pjFixedWork
            t.Work = workMinutes
        End If

        ' Assign Monteurs work resource
        Set rMonteurs = GetOrCreateWorkResourceCached("Monteurs", resourceCache)
        rMonteurs.MaxUnits = 10
        If workMinutes > 0 Then
            Dim nbPers As Long
            nbPers = IIf(IsNumeric(pers) And CDbl(pers) > 0, CLng(pers), 1)

            Set a = t.Assignments.Add(ResourceID:=rMonteurs.ID)
            a.Work = workMinutes
            a.Units = nbPers
            a.Work = workMinutes
            a.WorkContour = pjFlat
            dateDebutMonteurs = a.Start
            dateFinMonteurs = a.Finish
            hasMonteursAssignment = True
            ' copy tags to assignment
            a.Text1 = tranche: a.Text2 = zone: a.Text3 = sousZone: a.Text4 = typ
            a.Text5 = entreprise: a.Text6 = niveau: a.Text7 = onduleur: a.Text8 = ptr
        End If

' ==============================
' ONDULEUR MATERIAL RESOURCE (Column K)
' ==============================
Dim qteOnd As Double

If IsNumeric(niveau) Then
    qteOnd = CDbl(niveau)
    
    If qteOnd > 0 Then
        Dim rOnd As Resource
        Dim onduleurName As String
        
        onduleurName = "Onduleurs " & nom  ' Task name
        
        ' Create or get material resource
        Set rOnd = GetOrCreateMaterialResourceCached(onduleurName, resourceCache)
        
        ' Assign resource to task
        Set a = t.Assignments.Add(ResourceID:=rOnd.ID)
        a.Units = qteOnd
        a.WorkContour = pjFlat
        
        ' Sync dates with Monteurs if present
        If hasMonteursAssignment Then
            a.Start = dateDebutMonteurs
            a.Finish = dateFinMonteurs
        End If
        
        ' Copy tags Text1-Text8
        a.Text1 = tranche
        a.Text2 = zone
        a.Text3 = sousZone
        a.Text4 = typ
        a.Text5 = entreprise
        a.Text6 = qteOnd   ' numeric quantity
        a.Text7 = onduleur
        a.Text8 = ptr
    End If
End If



        ' Quality CQ handling (simplified to match original logic)
        Dim isOmx As Boolean
        isOmx = (UCase$(entreprise) = "OMX" Or UCase$(entreprise) = "OMEXOM")
        If qualite = "CQ" Then
            Dim rCQMat As Resource
            Set rCQMat = GetOrCreateMaterialResourceCached("CQ", resourceCache)
            If isOmx Then
                Set a = t.Assignments.Add(ResourceID:=rCQMat.ID)
                a.Units = 1
                a.WorkContour = pjFlat
                If hasMonteursAssignment Then a.Start = dateDebutMonteurs: a.Finish = dateFinMonteurs
                a.Text1 = tranche: a.Text2 = zone: a.Text3 = sousZone: a.Text4 = typ
                a.Text5 = entreprise: a.Text6 = niveau: a.Text7 = onduleur: a.Text8 = ptr
            Else
                Set tCQ = pjProj.Tasks.Add("Contrôle Qualité - " & nom)
                tCQ.Manual = False: tCQ.Calendar = pjProj.BaseCalendars("Standard"): tCQ.LevelingCanSplit = False
                tCQ.Text1 = tranche: tCQ.Text2 = zone: tCQ.Text3 = sousZone: tCQ.Text4 = "CQ"
                tCQ.Text5 = "OMEXOM": tCQ.Text6 = niveau: tCQ.Text7 = onduleur: tCQ.Text8 = ptr
                Set a = tCQ.Assignments.Add(ResourceID:=rCQMat.ID)
                a.Units = 1: a.WorkContour = pjFlat
                ' Link start to start +1d
                On Error Resume Next
                t.LinkSuccessors tCQ, pjStartToStart, "1d"
                On Error GoTo 0
            End If
        ElseIf qualite = "TACHE" Or qualite = "TÂCHE" Then
            Set tCQ = pjProj.Tasks.Add("Contrôle Qualité - " & nom)
            tCQ.Manual = False: tCQ.Calendar = pjProj.BaseCalendars("Standard"): tCQ.LevelingCanSplit = False
            tCQ.Text1 = tranche: tCQ.Text2 = zone: tCQ.Text3 = sousZone: tCQ.Text4 = "CQ"
            tCQ.Text5 = "OMEXOM": tCQ.Text6 = niveau: tCQ.Text7 = onduleur: tCQ.Text8 = ptr
            Dim rCQMat2 As Resource
            Set rCQMat2 = GetOrCreateMaterialResourceCached("CQ", resourceCache)
            Set a = tCQ.Assignments.Add(ResourceID:=rCQMat2.ID)
            a.Units = 1: a.WorkContour = pjFlat
            On Error Resume Next
            t.LinkSuccessors tCQ, pjStartToStart, "1d"
            On Error GoTo 0
        End If

NextRow:
        ' continue loop
    Next i
    
    ' ==============================
    ' RECALCULATE PROJECT
    ' ==============================
    
    pjApp.CalculateProject
    
    ' ==============================
' SAVE PROJECT AS MPP (A2 + Timestamp)
' ==============================

Dim baseName As String
Dim timestamp As String
Dim folderPath As String
Dim savePath As String

' Get project name from Excel A2
baseName = Trim(xlSheet.Range("A2").Value)

If baseName = "" Then
    MsgBox "Cell A2 is empty. Cannot generate project name.", vbExclamation
    GoTo CleanExit
End If

' Remove invalid filename characters
baseName = Replace(baseName, "/", "-")
baseName = Replace(baseName, "\", "-")
baseName = Replace(baseName, ":", "-")
baseName = Replace(baseName, "*", "-")
baseName = Replace(baseName, "?", "-")
baseName = Replace(baseName, """", "-")
baseName = Replace(baseName, "<", "-")
baseName = Replace(baseName, ">", "-")
baseName = Replace(baseName, "|", "-")

' Create timestamp
timestamp = Format(Now, "yyyymmdd_hhnnss")

' Extract folder from Excel path
folderPath = Left(fichierExcel, InStrRev(fichierExcel, "\"))

' Final path
savePath = folderPath & baseName & "_" & timestamp & ".mpp"

' Save project
pjProj.SaveAs Name:=savePath

    
    ' ==============================
    ' REBUILD PLANO MENU
    ' ==============================
    
    CreatePlanoMenu
    
    MsgBox "Projet créé avec succès !" & vbCrLf & savePath, vbInformation

CleanExit:

    On Error Resume Next
    
    If Not xlBook Is Nothing Then xlBook.Close False
    If Not xlApp Is Nothing Then xlApp.Quit
    
    Set xlSheet = Nothing
    Set xlBook = Nothing
    Set xlApp = Nothing
    
    pjApp.ScreenUpdating = oldScreenUpdating
    pjApp.DisplayAlerts = True

End Sub

' ======= HELPERS =======
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

Function GetOrCreateWorkResourceCached(nom As String, ByRef cache As Object) As Resource

    Dim r As Resource

    ' Try to get resource directly from project
    On Error Resume Next
    Set r = ActiveProject.Resources(nom)
    On Error GoTo 0

    ' If not found, create it
    If r Is Nothing Then
        Set r = ActiveProject.Resources.Add(nom)
        r.Type = pjResourceTypeWork
    End If

    ' Store only the NAME in cache (not object)
    If Not cache.Exists(nom) Then
        cache.Add nom, True
    End If

    Set GetOrCreateWorkResourceCached = r

End Function




Function GetOrCreateMaterialResourceCached(nom As String, ByRef cache As Object) As Resource

    Dim r As Resource
    Dim cacheKey As String
    
    cacheKey = "MAT_" & nom

    On Error Resume Next
    Set r = ActiveProject.Resources(nom)
    On Error GoTo 0

    If r Is Nothing Then
        Set r = ActiveProject.Resources.Add(nom)
        r.Type = pjResourceTypeMaterial
    End If

    If Not cache.Exists(cacheKey) Then
        cache.Add cacheKey, True
    End If

    Set GetOrCreateMaterialResourceCached = r

End Function

