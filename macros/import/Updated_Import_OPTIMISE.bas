Attribute VB_Name = "Module1"
Option Explicit

Sub Import_Taches_Simples_AvecTitre()

    ' ==============================
    ' DECLARATIONS
    ' ==============================
    Dim xlApp As Object, xlBook As Object, xlSheet As Object
    Dim pjApp As MSProject.Application, pjProj As MSProject.Project
    Dim dataArr As Variant
    Dim resourceCache As Object
    Dim oldScreenUpdating As Boolean
    Dim oldCalculation As Boolean
    Dim fichierExcel As String
    Dim lastRow As Long
    Dim i As Long
    Dim t As Task, tCQ As Task, tGroup As Task, tRoot As Task
    Dim a As Assignment
    Dim listSep As String  ' <-- System list separator for resource name sanitization

    ' Log file
    Dim fso As Object, logStream As Object
    Dim logFile As String

    ' ==============================
    ' USE CURRENT PROJECT
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
    oldCalculation = pjApp.Calculation
    pjApp.Calculation = False
    On Error GoTo 0

    ' ==============================
    ' SELECT EXCEL FILE
    ' ==============================
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = False
    fichierExcel = xlApp.GetOpenFilename("Fichiers Excel (*.xlsx;*.xls), *.xlsx;*.xls")
    If fichierExcel = "False" Then
        MsgBox "Import annulé.", vbInformation
        GoTo CleanExit
    End If

    ' Get Windows list separator from Excel (xlListSeparator = 5)
    On Error Resume Next
    listSep = xlApp.Application.International(5)
    If Err.Number <> 0 Or listSep = "" Then listSep = ","
    Err.Clear
    On Error GoTo 0

    Set xlBook = xlApp.Workbooks.Open(FileName:=fichierExcel, ReadOnly:=True, UpdateLinks:=False)
    Set xlSheet = xlBook.Sheets(1)

    ' ==============================
    ' LOAD DATA INTO MEMORY ARRAY (A1:M<lastRow>)
    ' ==============================
    lastRow = xlSheet.Cells(xlSheet.Rows.Count, 1).End(-4162).Row ' xlUp
    If lastRow < 2 Then
        MsgBox "Fichier Excel vide.", vbExclamation
        GoTo CleanExit
    End If

    ' Include header row so that A2 = dataArr(2,1)
    dataArr = xlSheet.Range("A1:M" & lastRow).Value

    ' ==============================
    ' RESOURCE CACHE (OBJECT DICTIONARY)
    ' ==============================
    Set resourceCache = CreateObject("Scripting.Dictionary")

    Dim r As Resource, safe As String
    For Each r In pjProj.Resources
        If Not r Is Nothing Then
            ' Cache raw name
            If Not resourceCache.Exists(r.Name) Then resourceCache.Add r.Name, r
            ' Cache sanitized name (so helpers can hit the cache)
            safe = CleanResourceName(r.Name, listSep)
            If Not resourceCache.Exists(safe) Then resourceCache.Add safe, r
            ' Also pre-fill MAT_ key for materials (same name)
            If r.Type = pjResourceTypeMaterial Then
                If Not resourceCache.Exists("MAT_" & r.Name) Then resourceCache.Add "MAT_" & r.Name, r
                If Not resourceCache.Exists("MAT_" & safe) Then resourceCache.Add "MAT_" & safe, r
            End If
        End If
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
    pjApp.CustomFieldRename pjCustomTaskNumber5, "Qté Onduleurs"
    pjApp.CustomFieldRename pjCustomTaskText8, "PTR"

    ' ==============================
    ' ROOT TASK FROM A2
    ' ==============================
    Dim projTitle As String
    projTitle = Trim$(CStr(dataArr(2, 1)))
    If projTitle <> "" Then
        On Error Resume Next
        Set tRoot = pjProj.Tasks.Add(Name:=projTitle, Before:=1)
        On Error GoTo 0
        If Not tRoot Is Nothing Then
            tRoot.Manual = False
            On Error Resume Next
            tRoot.Calendar = pjProj.BaseCalendars("Standard")
            On Error GoTo 0
            Set tGroup = tRoot
        End If
    End If

    ' ==============================
    ' LOG FILE SETUP
    ' ==============================
    Set fso = CreateObject("Scripting.FileSystemObject")
    logFile = Replace(fichierExcel, ".xlsx", "_import_log.txt")
    logFile = Replace(logFile, ".xls", "_import_log.txt")
    Set logStream = fso.CreateTextFile(logFile, True)
    logStream.WriteLine "===== DEBUT IMPORT - " & Now & " ====="
    logStream.WriteLine "Fichier source: " & fichierExcel
    logStream.WriteLine "Nombre de lignes: " & lastRow
    logStream.WriteLine "[OPTIMISE] Lecture Array + Cache Ressources + ScreenUpdating OFF"
    logStream.WriteLine ""

    ' ==============================
    ' MAIN LOOP: ROWS 3..lastRow
    ' ==============================
    Dim nom As String, qte As Variant, pers As Variant, h As Variant
    Dim zone As String, sousZone As String, tranche As String, typ As String, entreprise As String
    Dim qualite As String, niveau As String, ptr As String
    Dim qteOnduleurs As Double
    Dim dateDebutMonteurs As Date, dateFinMonteurs As Date
    Dim hasMonteursAssignment As Boolean

    For i = 3 To lastRow

        ' ---- Read row i from dataArr (A..M = 1..13)
        nom = Trim$(CStr(dataArr(i, 1)))                ' A
        qte = dataArr(i, 2)                             ' B
        pers = dataArr(i, 3)                            ' C
        h = dataArr(i, 4)                               ' D
        zone = Trim$(CStr(dataArr(i, 5)))               ' E
        sousZone = Trim$(CStr(dataArr(i, 6)))           ' F
        tranche = Trim$(CStr(dataArr(i, 7)))            ' G
        typ = Trim$(CStr(dataArr(i, 8)))                ' H
        entreprise = Trim$(CStr(dataArr(i, 9)))         ' I
        qualite = UCase$(Trim$(CStr(dataArr(i, 10))))   ' J: CQ / TACHE / vide
        niveau = UCase$(Trim$(CStr(dataArr(i, 11))))    ' K: SZ / OND / vide

        ' L: Quantité d'onduleurs (Double)
        On Error Resume Next
        qteOnduleurs = 0
        If Len(Trim$(CStr(dataArr(i, 12)))) > 0 Then
            qteOnduleurs = CDbl(dataArr(i, 12))
        End If
        If Err.Number <> 0 Then qteOnduleurs = 0
        Err.Clear

        ' M: PTR (optionnel)
        ptr = ""
        If Not IsError(dataArr(i, 13)) Then
            ptr = Trim$(CStr(dataArr(i, 13)))
        End If
        On Error GoTo 0

        hasMonteursAssignment = False

        ' Skip blank name
        If nom = "" Then GoTo NextRow

        ' ---- TITLE detection (no quantity and no hours) -> create group at level 2
        If IsEmptyOrZero(qte) And IsEmptyOrZero(h) Then
            On Error Resume Next
            Set tGroup = pjProj.Tasks.Add(nom)
            If Err.Number <> 0 Then
                logStream.WriteLine "[WARN] Impossible de créer le TITRE pour: " & nom & " (Err " & Err.Number & ": " & Err.Description & ")"
                Err.Clear
                On Error GoTo 0
                GoTo NextRow
            End If
            On Error GoTo 0

            tGroup.Manual = False

            ' Ensure level 2 (flat under project/root)
            Call EnsureOutlineLevel(tGroup, 2)

            ' Tag recap to allow filtering at higher levels
            tGroup.Text1 = tranche
            tGroup.Text2 = zone
            tGroup.Text3 = sousZone
            tGroup.Text4 = typ
            tGroup.Text5 = entreprise
            tGroup.Text6 = niveau
            tGroup.Number5 = qteOnduleurs
            tGroup.Text8 = ptr

            logStream.WriteLine "TITRE: " & tGroup.Name & " -> Niveau " & tGroup.OutlineLevel
            GoTo NextRow
        End If

        ' ---- TASK creation
        On Error Resume Next
        Set t = pjProj.Tasks.Add(nom)
        If Err.Number <> 0 Or t Is Nothing Then
            logStream.WriteLine "[ERREUR] Tasks.Add(): " & nom & " | " & Err.Number & " - " & Err.Description
            Err.Clear
            On Error GoTo 0
            GoTo NextRow
        End If
        On Error GoTo 0

        t.Manual = False
        On Error Resume Next
        t.Calendar = pjProj.BaseCalendars("Standard")
        On Error GoTo 0
        t.LevelingCanSplit = False

        ' Determine target outline level
        Dim targetLevel As Integer
        If niveau = "OND" Then
            targetLevel = 4
        ElseIf niveau = "SZ" Then
            targetLevel = 3
        Else
            If Not tGroup Is Nothing Then
                targetLevel = tGroup.OutlineLevel + 1
            Else
                targetLevel = 3
            End If
        End If
        Call EnsureOutlineLevel(t, targetLevel)

        ' Tag task
        t.Text1 = tranche
        t.Text2 = zone
        t.Text3 = sousZone
        t.Text4 = typ
        t.Text5 = entreprise
        t.Text6 = niveau
        t.Number5 = qteOnduleurs
        t.Text8 = ptr

        ' ---- Set Work on task first (minutes)
        Dim workMinutes As Long
        workMinutes = 0
        If IsNumeric(h) And CDbl(h) > 0 Then
            workMinutes = CLng(CDbl(h) * 60#)
            t.Type = pjFixedWork
            t.Work = workMinutes
        End If

        ' ---- Assign Monteurs (work resource) if hours present
        If workMinutes > 0 Then
            Dim nbPers As Long
            nbPers = IIf(IsNumeric(pers) And CDbl(pers) > 0, CLng(pers), 1)

            Dim rMonteurs As Resource
            Set rMonteurs = GetOrCreateWorkResourceCached("Monteurs", resourceCache, listSep)
            rMonteurs.MaxUnits = 10 ' 1000%

            Set a = t.Assignments.Add(ResourceID:=rMonteurs.ID)
            a.Work = workMinutes         ' Step 1: Work first
            a.Units = nbPers             ' Step 2: Units
            a.Work = workMinutes         ' Step 3: Force Work again
            a.WorkContour = pjFlat       ' Step 4: Flat contour

            dateDebutMonteurs = a.Start
            dateFinMonteurs = a.Finish
            hasMonteursAssignment = True

            ' Copy tags
            a.Text1 = tranche
            a.Text2 = zone
            a.Text3 = sousZone
            a.Text4 = typ
            a.Text5 = entreprise
            a.Text6 = niveau
            a.Number5 = qteOnduleurs
            a.Text8 = ptr
        End If

        ' ---- Material quantity (column B)
        If IsNumeric(qte) And CDbl(qte) > 0 Then
            Dim rMat As Resource
            Dim nomRessource As String

            If niveau = "OND" And Not tGroup Is Nothing Then
                nomRessource = tGroup.Name    ' aggregate by group
            Else
                nomRessource = t.Name         ' task name
            End If

            Set rMat = GetOrCreateMaterialResourceCached(nomRessource, resourceCache, listSep)
            Set a = t.Assignments.Add(ResourceID:=rMat.ID)
            a.Units = CDbl(qte)
            a.WorkContour = pjFlat

            If hasMonteursAssignment Then
                a.Start = dateDebutMonteurs
                a.Finish = dateFinMonteurs
            End If

            ' Copy tags
            a.Text1 = tranche
            a.Text2 = zone
            a.Text3 = sousZone
            a.Text4 = typ
            a.Text5 = entreprise
            a.Text6 = niveau
            a.Number5 = qteOnduleurs
            a.Text8 = ptr
        End If

        ' ---- ONDULEUR material (column L)
        If qteOnduleurs > 0 Then
            Dim rOnduleur As Resource
            Set rOnduleur = GetOrCreateMaterialResourceCached("ONDULEUR", resourceCache, listSep)

            Set a = t.Assignments.Add(ResourceID:=rOnduleur.ID)
            a.Units = qteOnduleurs
            a.WorkContour = pjFlat

            If hasMonteursAssignment Then
                a.Start = dateDebutMonteurs
                a.Finish = dateFinMonteurs
            End If

            ' Copy tags
            a.Text1 = tranche
            a.Text2 = zone
            a.Text3 = sousZone
            a.Text4 = typ
            a.Text5 = entreprise
            a.Text6 = niveau
            a.Number5 = qteOnduleurs
            a.Text8 = ptr
        End If

        ' ---- QUALITE hybrid logic
        Dim isOmx As Boolean
        isOmx = (UCase$(entreprise) = "OMX" Or UCase$(entreprise) = "OMEXOM")

        If qualite = "CQ" Then
            Dim rCQMat As Resource
            Set rCQMat = GetOrCreateMaterialResourceCached("CQ", resourceCache, listSep)

            If isOmx Then
                ' Inline CQ as material on the task
                Set a = t.Assignments.Add(ResourceID:=rCQMat.ID)
                a.Units = 1
                a.WorkContour = pjFlat
                If hasMonteursAssignment Then a.Start = dateDebutMonteurs: a.Finish = dateFinMonteurs

                a.Text1 = tranche
                a.Text2 = zone
                a.Text3 = sousZone
                a.Text4 = typ
                a.Text5 = entreprise
                a.Text6 = niveau
                a.Number5 = qteOnduleurs
                a.Text8 = ptr
            Else
                ' Separate CQ task with SS+1d link
                Set tCQ = pjProj.Tasks.Add("Contrôle Qualité - " & nom)
                tCQ.Manual = False
                On Error Resume Next
                tCQ.Calendar = pjProj.BaseCalendars("Standard")
                On Error GoTo 0
                tCQ.LevelingCanSplit = False

                ' Align outline close to parent task
                Call EnsureOutlineLevel(tCQ, t.OutlineLevel)

                ' Tags
                tCQ.Text1 = tranche
                tCQ.Text2 = zone
                tCQ.Text3 = sousZone
                tCQ.Text4 = "CQ"
                tCQ.Text5 = "OMEXOM"
                tCQ.Text6 = niveau
                tCQ.Number5 = qteOnduleurs
                tCQ.Text8 = ptr

                ' CQ material assignment
                Set a = tCQ.Assignments.Add(ResourceID:=rCQMat.ID)
                a.Units = 1
                a.WorkContour = pjFlat

                ' Link Start-to-Start +1d
                On Error Resume Next
                t.LinkSuccessors tCQ, pjStartToStart, "1d"
                On Error GoTo 0
            End If

        ElseIf qualite = "TACHE" Or qualite = "TÂCHE" Then
            ' Force separate CQ task
            Dim rCQMat2 As Resource
            Set rCQMat2 = GetOrCreateMaterialResourceCached("CQ", resourceCache, listSep)

            Set tCQ = pjProj.Tasks.Add("Contrôle Qualité - " & nom)
            tCQ.Manual = False
            On Error Resume Next
            tCQ.Calendar = pjProj.BaseCalendars("Standard")
            On Error GoTo 0
            tCQ.LevelingCanSplit = False

            Call EnsureOutlineLevel(tCQ, t.OutlineLevel)

            tCQ.Text1 = tranche
            tCQ.Text2 = zone
            tCQ.Text3 = sousZone
            tCQ.Text4 = "CQ"
            tCQ.Text5 = "OMEXOM"
            tCQ.Text6 = niveau
            tCQ.Number5 = qteOnduleurs
            tCQ.Text8 = ptr

            Set a = tCQ.Assignments.Add(ResourceID:=rCQMat2.ID)
            a.Units = 1
            a.WorkContour = pjFlat

            On Error Resume Next
            t.LinkSuccessors tCQ, pjStartToStart, "1d"
            On Error GoTo 0
        End If

NextRow:
        ' continue
    Next i

    ' ==============================
    ' FINALIZE: CALCULATE & SAVE
    ' ==============================
    On Error Resume Next
    pjApp.Calculation = True
    pjProj.Calculate
    pjApp.CalculateAll
    On Error GoTo 0

    ' Save as (A2 + timestamp) in same folder as Excel
    Dim baseName As String, timestamp As String
    Dim folderPath As String, savePath As String

    baseName = Trim$(CStr(xlSheet.Range("A2").Value))
    If baseName = "" Then baseName = projTitle
    If baseName = "" Then baseName = "Projet"

    baseName = SanitizeFileName(baseName)
    timestamp = Format(Now, "yyyymmdd_hhnnss")
    folderPath = Left$(fichierExcel, InStrRev(fichierExcel, "\"))
    savePath = folderPath & baseName & "_" & timestamp & ".mpp"

    pjProj.SaveAs Name:=savePath

    ' UI refresh
    On Error Resume Next
    pjApp.ScreenRefresh
    On Error GoTo 0

    logStream.WriteLine ""
    logStream.WriteLine "===== FIN IMPORT - " & Now & " ====="
    logStream.Close
    Set logStream = Nothing
    Set fso = Nothing

    ' ==============================
    ' REBUILD PLANO MENU
    ' ==============================
    On Error Resume Next
    CreatePlanoMenu
    On Error GoTo 0

    MsgBox "Import terminé: tâches, ressources, tags (Zone/Sous-zone/Tranche/Type/Entreprise/Niveau/Qté Onduleurs/PTR) et Qualité hybride." & vbCrLf & vbCrLf & _
           "Projet sauvegardé: " & savePath & vbCrLf & "Fichier log: " & logFile, vbInformation

CleanExit:
    On Error Resume Next
    If Not xlBook Is Nothing Then xlBook.Close False
    If Not xlApp Is Nothing Then xlApp.Quit
    Set xlSheet = Nothing: Set xlBook = Nothing: Set xlApp = Nothing

    pjApp.ScreenUpdating = oldScreenUpdating
    pjApp.Calculation = oldCalculation
    pjApp.DisplayAlerts = True
End Sub

' ==============================
' HELPERS
' ==============================

Private Sub EnsureOutlineLevel(ByVal tsk As Task, ByVal targetLevel As Integer)
    On Error Resume Next
    If tsk Is Nothing Then Exit Sub
    If targetLevel < 1 Then targetLevel = 1
    If targetLevel > 9 Then targetLevel = 9

    If tsk.OutlineLevel < targetLevel Then
        Do While tsk.OutlineLevel < targetLevel And tsk.OutlineLevel < 9 And Not tsk.Summary
            tsk.OutlineIndent
            If Err.Number <> 0 Then Exit Do
        Loop
    ElseIf tsk.OutlineLevel > targetLevel Then
        Do While tsk.OutlineLevel > targetLevel And tsk.OutlineLevel > 1 And Not tsk.Summary
            tsk.OutlineOutdent
            If Err.Number <> 0 Then Exit Do
        Loop
    End If
    Err.Clear
    On Error GoTo 0
End Sub

Private Function SanitizeFileName(ByVal s As String) As String
    s = Replace(s, "/", "-")
    s = Replace(s, "\", "-")
    s = Replace(s, ":", "-")
    s = Replace(s, "*", "-")
    s = Replace(s, "?", "-")
    s = Replace(s, """", "-")
    s = Replace(s, "<", "-")
    s = Replace(s, ">", "-")
    s = Replace(s, "|", "-")
    SanitizeFileName = s
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

' ---- Diagnostics from .vb (kept for parity / 1101 debugging)
Private Function AnalyzeStringCharacters(text As String) As String
    Dim result As String
    Dim i As Integer
    Dim ch As String
    Dim charCode As Long
    Dim charName As String

    result = ""
    For i = 1 To Len(text)
        ch = Mid$(text, i, 1)
        charCode = AscW(ch)

        Select Case charCode
            Case 9: charName = "TAB"
            Case 10: charName = "LF (Line Feed)"
            Case 13: charName = "CR (Carriage Return)"
            Case 32: charName = "SPACE"
            Case 160: charName = "NBSP (Non-Breaking Space)"
            Case 0 To 31: charName = "CTRL"
            Case 127 To 159: charName = "CTRL étendu"
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

Private Function IsInvisibleOnlyString(text As String) As Boolean
    Dim i As Integer
    Dim ch As String
    Dim charCode As Long
    Dim hasVisibleChar As Boolean

    hasVisibleChar = False

    For i = 1 To Len(text)
        ch = Mid$(text, i, 1)
        charCode = AscW(ch)

        If (charCode >= 33 And charCode <= 126) Or (charCode > 160) Then
            hasVisibleChar = True
            Exit For
        End If
    Next i

    IsInvisibleOnlyString = Not hasVisibleChar
End Function

Private Function IsNumericPattern(text As String) As Boolean
    Dim i As Integer, ch As String
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

' Copy tags helper (not strictly required since we inline copy, but kept)
Private Sub CopyTaskTagsToAssignment(ByVal tSource As Task, ByVal a As Assignment)
    On Error GoTo ErrHandler
    DoEvents
    a.Text1 = tSource.Text1
    a.Text2 = tSource.Text2
    a.Text3 = tSource.Text3
    a.Text4 = tSource.Text4
    a.Text5 = tSource.Text5
    a.Text6 = tSource.Text6
    a.Number5 = tSource.Number5
    a.Text8 = tSource.Text8
    Exit Sub
ErrHandler:
    Debug.Print "ERREUR CopyTaskTagsToAssignment: " & Err.Description & " (Tâche: " & tSource.Name & ")"
    Resume Next
End Sub

' ==============================
' Resource name sanitization
' ==============================

Private Function ReplaceMultipleSpaces(ByVal s As String) As String
    Do While InStr(1, s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    ReplaceMultipleSpaces = s
End Function

' Remove forbidden characters: [ ] and system list separator; normalize whitespace/control chars; enforce length
Private Function CleanResourceName(ByVal raw As String, Optional ByVal listSep As String = ",") As String
    Dim s As String, i As Long, ch As String, code As Long, out As String

    s = Trim$(raw)
    If s = "" Then
        CleanResourceName = "Ressource"
        Exit Function
    End If

    ' Normalize whitespace and NBSPs
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    s = Replace(s, vbTab, " ")
    s = Replace(s, Chr$(160), " ")

    ' Strip forbidden: [ ] and list separator
    s = Replace(s, "[", " ")
    s = Replace(s, "]", " ")
    If listSep <> "" Then s = Replace(s, listSep, " ")

    ' Remove control characters; keep visible chars
    out = ""
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        code = AscW(ch)
        If (code >= 32 And code <> 127 And Not (code >= 128 And code <= 159)) Then
            out = out & ch
        Else
            out = out & " "
        End If
    Next i

    ' Collapse spaces and limit length
    out = Trim$(ReplaceMultipleSpaces(out))
    If Len(out) > 255 Then out = Left$(out, 255)
    If out = "" Then out = "Ressource"

    CleanResourceName = out
End Function

' ---- Resource helpers with sanitization and cache

Function GetOrCreateWorkResourceCached(nom As String, ByRef cache As Object, Optional ByVal listSep As String = ",") As Resource
    Dim safeName As String
    safeName = CleanResourceName(nom, listSep)

    ' Cache hit?
    If cache.Exists(safeName) Then
        Set GetOrCreateWorkResourceCached = cache(safeName)
        Exit Function
    End If

    Dim r As Resource
    On Error Resume Next
    Set r = ActiveProject.Resources(safeName)
    On Error GoTo 0

    If r Is Nothing Then
        On Error Resume Next
        Set r = ActiveProject.Resources.Add(safeName)
        If Err.Number <> 0 Then
            ' Try with suffix to avoid collision/invalid
            Err.Clear
            Set r = ActiveProject.Resources.Add(safeName & " (W)")
        End If
        On Error GoTo 0

        If Not r Is Nothing Then
            On Error Resume Next
            r.Type = pjResourceTypeWork
            On Error GoTo 0
        Else
            ' Final fallback (rare)
            Set r = ActiveProject.Resources.Add("Ressource Travail")
            On Error Resume Next: r.Type = pjResourceTypeWork: On Error GoTo 0
        End If
    End If

    If Not cache.Exists(safeName) Then cache.Add safeName, r
    Set GetOrCreateWorkResourceCached = r
End Function

Function GetOrCreateMaterialResourceCached(nom As String, ByRef cache As Object, Optional ByVal listSep As String = ",") As Resource
    Dim safeName As String, cacheKey As String
    safeName = CleanResourceName(nom, listSep)
    cacheKey = "MAT_" & safeName

    If cache.Exists(cacheKey) Then
        Set GetOrCreateMaterialResourceCached = cache(cacheKey)
        Exit Function
    End If

    Dim r As Resource
    On Error Resume Next
    Set r = ActiveProject.Resources(safeName)
    On Error GoTo 0

    If r Is Nothing Then
        On Error Resume Next
        Set r = ActiveProject.Resources.Add(safeName)
        If Err.Number <> 0 Then
            ' Try with suffix
            Err.Clear
            Set r = ActiveProject.Resources.Add(safeName & " (M)")
        End If
        On Error GoTo 0

        If Not r Is Nothing Then
            On Error Resume Next
            r.Type = pjResourceTypeMaterial
            On Error GoTo 0
        Else
            ' Final fallback
            Set r = ActiveProject.Resources.Add("Ressource Matériau")
            On Error Resume Next: r.Type = pjResourceTypeMaterial: On Error GoTo 0
        End If
    Else
        ' If found but wrong type, create distinct material resource
        If r.Type <> pjResourceTypeMaterial Then
            On Error Resume Next
            Set r = ActiveProject.Resources.Add(safeName & " (M)")
            Err.Clear
            r.Type = pjResourceTypeMaterial
            On Error GoTo 0
        End If
    End If

    If Not cache.Exists(cacheKey) Then cache.Add cacheKey, r
    Set GetOrCreateMaterialResourceCached = r
End Function
