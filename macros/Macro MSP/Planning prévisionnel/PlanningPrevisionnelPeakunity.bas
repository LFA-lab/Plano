Sub Export_Plan_Et_UnitesDePointe()
    Dim xlApp As Object
    Dim xlWorkbook As Object

    Application.StatusBar = "Export en cours, initialisation Excel..."
    Call VCPlanningjour(xlApp, xlWorkbook)
    Application.StatusBar = "G�n�ration du planning pr�visionnel, veuillez patienter..."
    DoEvents

    Application.StatusBar = "G�n�ration du planning pr�visionnel, Injection des unit�s de pointe..."
    Call Injecter_UnitesDePointe_Dans_Planning(xlApp, xlWorkbook)
    Application.StatusBar = False

    If Not xlApp Is Nothing Then
        xlApp.Visible = True
        xlApp.ScreenUpdating = True
    End If
End Sub

' ==============================
' G�N�RATION PLANNING ET OUTILS
' ==============================
Sub VCPlanningjour(ByRef xlApp As Object, ByRef xlWorkbook As Object)
    Dim projDoc As Object
    Dim wsJours As Object
    Dim wsSemaines As Object

    On Error GoTo ErrorHandler

    If ActiveProject Is Nothing Then
        MsgBox "Aucun projet MS Project n'est ouvert!", vbCritical
        Exit Sub
    End If

    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = False
    xlApp.ScreenUpdating = False
    Set xlWorkbook = xlApp.Workbooks.Add
    Set projDoc = ActiveProject

    Dim allTasks As Variant
    Dim taskCount As Integer
    Application.StatusBar = "G�n�ration du planning pr�visionnel, Collecte des t�ches MS Project..."
    taskCount = CollectTasksData(projDoc, allTasks)
    DoEvents

    If taskCount = 0 Then
        MsgBox "Aucune t�che trouv�e dans le projet!", vbExclamation
        GoTo Cleanup
    End If

    Application.StatusBar = "G�n�ration du planning pr�visionnel, G�n�ration de la vue Jours..."
    Set wsJours = xlWorkbook.Worksheets(1)
    Call GenerateDayView(wsJours, allTasks, taskCount, projDoc)
    wsJours.Name = "Jours"
    InsererLogoOmexom wsJours
    DoEvents

    Application.StatusBar = "G�n�ration du planning pr�visionnel, G�n�ration de la vue Semaines..."
    Set wsSemaines = xlWorkbook.Worksheets.Add(After:=wsJours)
    Call GenerateWeekView(wsSemaines, allTasks, taskCount, projDoc)
    wsSemaines.Name = "Semaines"
    InsererLogoOmexom wsSemaines
    DoEvents

Cleanup:
    Exit Sub

ErrorHandler:
    MsgBox "Erreur : " & Err.Number & " - " & Err.Description, vbCritical
    Exit Sub
End Sub

Function CollectTasksData(projDoc As Object, ByRef allTasks As Variant) As Integer
    Dim task As Object
    Dim taskCount As Integer
    Dim i As Integer

    taskCount = 0
    For Each task In projDoc.Tasks
        If Not task Is Nothing And Not task.Summary Then
            taskCount = taskCount + 1
        End If
    Next task

    If taskCount = 0 Then
        CollectTasksData = 0
        Exit Function
    End If

    ReDim allTasks(1 To taskCount, 1 To 8)
    ' 1:ID, 2:Name, 3:Start, 4:Finish, 5:BaselineStart, 6:BaselineFinish, 7:HasSchedule, 8:HasBaseline

    i = 1
    For Each task In projDoc.Tasks
        If Not task Is Nothing And Not task.Summary Then
            allTasks(i, 1) = task.ID
            allTasks(i, 2) = task.Name

            ' Attribution directe des dates
            On Error Resume Next
            allTasks(i, 3) = task.Start
            allTasks(i, 4) = task.Finish
            allTasks(i, 5) = task.BaselineStart
            allTasks(i, 6) = task.BaselineFinish
            On Error GoTo 0

            ' V�rifications robustes sur les dates
            allTasks(i, 7) = (IsDate(allTasks(i, 3)) And IsDate(allTasks(i, 4))) ' HasSchedule
            allTasks(i, 8) = (IsDate(allTasks(i, 5)) And IsDate(allTasks(i, 6))) ' HasBaseline

            i = i + 1
        End If
    Next task

    CollectTasksData = taskCount
End Function


Sub GenerateWeekView(ws As Object, allTasks As Variant, taskCount As Integer, projDoc As Object)
    Const MAX_WEEKS = 104
    Dim startDate As Date
    Dim weekHeaders() As String
    Dim weekDates() As Date
    Dim planningMatrix As Variant
    Dim colorRanges As Variant
    Dim colorCount As Integer

    startDate = projDoc.ProjectStart

    Call BuildWeekHeaders(ws, startDate, weekHeaders, weekDates, MAX_WEEKS)
    Call BuildPlanningMatrix(allTasks, taskCount, weekDates, planningMatrix, colorRanges, colorCount, True)
    Call DumpMatrixToSheet(ws, planningMatrix, UBound(weekHeaders))
    Call ApplyColorRanges(ws, colorRanges, colorCount)
    Call FormatWorksheetOptimized(ws, taskCount + 5, UBound(weekHeaders) + 2, "semaine")
End Sub

Sub GenerateDayView(ws As Object, allTasks As Variant, taskCount As Integer, projDoc As Object)
    Const MAX_DAYS = 180
    Dim startDate As Date
    Dim dayHeaders() As String
    Dim dayDates() As Date
    Dim planningMatrix As Variant
    Dim colorRanges As Variant
    Dim colorCount As Integer

    startDate = projDoc.ProjectStart

    Call BuildDayHeaders(ws, startDate, dayHeaders, dayDates, MAX_DAYS)
    Call BuildPlanningMatrix(allTasks, taskCount, dayDates, planningMatrix, colorRanges, colorCount, False)
    Call DumpMatrixToSheet(ws, planningMatrix, UBound(dayHeaders))
    Call ApplyColorRanges(ws, colorRanges, colorCount)
    Call FormatWorksheetOptimized(ws, taskCount + 5, UBound(dayHeaders) + 2, "jour")
End Sub

' Vue Semaine (ajout�e pour corriger la vue)
Sub BuildWeekHeaders(ws As Object, startDate As Date, ByRef weekHeaders() As String, ByRef weekDates() As Date, maxWeeks As Integer)
    Dim currentDate As Date
    Dim i As Integer
    Dim currentMonth As String
    Dim lastMonth As String
    Dim monthStartCol As Integer

    ' Trouver le d�but de semaine (lundi)
    currentDate = startDate
    Do While Weekday(currentDate, vbMonday) <> 1
        currentDate = currentDate - 1
    Loop

    ReDim weekHeaders(1 To maxWeeks)
    ReDim weekDates(1 To maxWeeks)

    For i = 1 To maxWeeks
        weekHeaders(i) = "S" & Format(currentDate, "ww")  ' Num�ro de semaine
        weekDates(i) = currentDate
        ws.Cells(100, i + 2).Value = currentDate         ' Pour correspondance ult�rieure
        currentDate = currentDate + 7                    ' Semaine suivante
    Next i

    ws.Cells(1, 1).Value = "PLANNING PR�VISIONNEL - VUE SEMAINE"
    ws.Cells(4, 1).Value = "N�"
    ws.Cells(4, 2).Value = "Nom de la t�che"

    lastMonth = ""
    monthStartCol = 3

    For i = 1 To maxWeeks
        currentMonth = Format(weekDates(i), "mmm-yy")
        ws.Cells(4, i + 2).Value = weekHeaders(i)
        If currentMonth <> lastMonth And lastMonth <> "" Then
            ws.Range(ws.Cells(2, monthStartCol), ws.Cells(2, i + 1)).Merge
            ws.Cells(2, monthStartCol).Value = lastMonth
            ws.Cells(2, monthStartCol).HorizontalAlignment = -4108
            monthStartCol = i + 2
        End If
        lastMonth = currentMonth
    Next i

    If maxWeeks > 0 Then
        ws.Range(ws.Cells(2, monthStartCol), ws.Cells(2, maxWeeks + 2)).Merge
        ws.Cells(2, monthStartCol).Value = lastMonth
        ws.Cells(2, monthStartCol).HorizontalAlignment = -4108
    End If

    ws.Rows(2).Font.Bold = True
    ws.Rows(4).Font.Bold = True
    ws.Rows(2).Interior.Color = RGB(200, 200, 200)
    ws.Rows(4).Interior.Color = RGB(230, 230, 230)
End Sub

Sub BuildDayHeaders(ws As Object, startDate As Date, ByRef dayHeaders() As String, ByRef dayDates() As Date, maxDays As Integer)
    Dim currentDate As Date
    Dim i As Integer
    Dim currentMonth As String
    Dim lastMonth As String
    Dim monthStartCol As Integer

    currentDate = startDate

    ReDim dayHeaders(1 To maxDays)
    ReDim dayDates(1 To maxDays)

    For i = 1 To maxDays
        dayHeaders(i) = Format(currentDate, "d")
        dayDates(i) = currentDate
        ws.Cells(100, i + 2).Value = currentDate  ' Stocke la date compl�te
        currentDate = currentDate + 1
    Next i

    ws.Cells(1, 1).Value = "PLANNING PR�VISIONNEL - VUE JOUR"
    ws.Cells(4, 1).Value = "N�"
    ws.Cells(4, 2).Value = "Nom de la t�che"

    lastMonth = ""
    monthStartCol = 3

    For i = 1 To maxDays
        currentMonth = Format(dayDates(i), "mmm-yy")
        ws.Cells(4, i + 2).Value = dayHeaders(i)
        If currentMonth <> lastMonth And lastMonth <> "" Then
            ws.Range(ws.Cells(2, monthStartCol), ws.Cells(2, i + 1)).Merge
            ws.Cells(2, monthStartCol).Value = lastMonth
            ws.Cells(2, monthStartCol).HorizontalAlignment = -4108
            monthStartCol = i + 2
        End If
        lastMonth = currentMonth
    Next i

    If maxDays > 0 Then
        ws.Range(ws.Cells(2, monthStartCol), ws.Cells(2, maxDays + 2)).Merge
        ws.Cells(2, monthStartCol).Value = lastMonth
        ws.Cells(2, monthStartCol).HorizontalAlignment = -4108
    End If

    ws.Rows(2).Font.Bold = True
    ws.Rows(4).Font.Bold = True
    ws.Rows(2).Interior.Color = RGB(200, 200, 200)
    ws.Rows(4).Interior.Color = RGB(230, 230, 230)
End Sub

Sub BuildPlanningMatrix(allTasks As Variant, taskCount As Integer, dates() As Date, _
                       ByRef planningMatrix As Variant, ByRef colorRanges As Variant, _
                       ByRef colorCount As Integer, isWeekView As Boolean)
    Dim i As Integer, j As Integer
    Dim maxCols As Integer
    Dim startCol As Integer, endCol As Integer
    Dim dStart As Date, dEnd As Date

    maxCols = UBound(dates)
    colorCount = 0

    ReDim planningMatrix(1 To taskCount, 1 To maxCols + 2)
    ReDim colorRanges(1 To taskCount * 4, 1 To 5)

    For i = 1 To taskCount
        planningMatrix(i, 1) = allTasks(i, 1)   ' ID
        planningMatrix(i, 2) = allTasks(i, 2)   ' Name

        For j = 3 To maxCols + 2
            planningMatrix(i, j) = ""
        Next j

        If Not allTasks(i, 7) Then
            planningMatrix(i, 3) = "Non planifi�"
        Else
            ' Baseline
            If allTasks(i, 8) Then
                If Not IsEmpty(allTasks(i, 5)) And Not IsEmpty(allTasks(i, 6)) And allTasks(i, 5) <> "" And allTasks(i, 6) <> "" Then
                    dStart = allTasks(i, 5)
                    dEnd = allTasks(i, 6)
                    Call ProcessDateRange(dStart, dEnd, dates, i, colorRanges, colorCount, "reference", isWeekView)
                End If
            End If
            ' R�el
            If Not IsEmpty(allTasks(i, 3)) And Not IsEmpty(allTasks(i, 4)) And allTasks(i, 3) <> "" And allTasks(i, 4) <> "" Then
                dStart = allTasks(i, 3)
                dEnd = allTasks(i, 4)
                Call ProcessDateRange(dStart, dEnd, dates, i, colorRanges, colorCount, "prevu", isWeekView)
            End If
        End If
    Next i
End Sub

Sub ProcessDateRange(startDate As Date, endDate As Date, dates() As Date, _
                    rowIndex As Integer, ByRef colorRanges As Variant, _
                    ByRef colorCount As Integer, colorType As String, isWeekView As Boolean)
    Dim i As Integer
    Dim rangeStart As Integer, rangeEnd As Integer
    Dim found As Boolean

    rangeStart = 0
    rangeEnd = 0

    For i = 1 To UBound(dates)
        If isWeekView Then
            If DateValue(dates(i)) + 6 >= DateValue(startDate) And DateValue(dates(i)) <= DateValue(endDate) Then
                If rangeStart = 0 Then rangeStart = i
                rangeEnd = i
                found = True
            End If
        Else
            If DateValue(dates(i)) >= DateValue(startDate) And DateValue(dates(i)) <= DateValue(endDate) Then
                If rangeStart = 0 Then rangeStart = i
                rangeEnd = i
                found = True
            End If
        End If
    Next i

    If found And rangeStart > 0 Then
        colorCount = colorCount + 1
        colorRanges(colorCount, 1) = rangeStart + 2  ' startCol
        colorRanges(colorCount, 2) = rangeEnd + 2    ' endCol
        colorRanges(colorCount, 3) = rowIndex + 4    ' rowIndex
        colorRanges(colorCount, 4) = colorType       ' colorType
    End If
End Sub

Sub DumpMatrixToSheet(ws As Object, planningMatrix As Variant, maxCols As Integer)
    Dim lastRow As Integer, lastCol As Integer

    lastRow = UBound(planningMatrix, 1)
    lastCol = maxCols + 2
    ws.Range(ws.Cells(5, 1), ws.Cells(lastRow + 5 - 1, lastCol)).Value = planningMatrix
End Sub

Sub ApplyColorRanges(ws As Object, colorRanges As Variant, colorCount As Integer)
    Dim i As Integer
    Dim targetRange As Object
    Dim baseColor As Long

    baseColor = RGB(255, 153, 51)

    For i = 1 To colorCount
        Set targetRange = ws.Range(ws.Cells(colorRanges(i, 3), colorRanges(i, 1)), ws.Cells(colorRanges(i, 3), colorRanges(i, 2)))
        targetRange.Interior.Color = baseColor
        If colorRanges(i, 4) = "reference" Then
            targetRange.Interior.TintAndShade = 0.8
        ElseIf colorRanges(i, 4) = "prevu" Then
            targetRange.Interior.TintAndShade = 0.2
        End If
    Next i
End Sub

Sub FormatWorksheetOptimized(ws As Object, lastRow As Integer, lastCol As Integer, viewType As String)
    With ws.Range("A1:F1")
        .Merge
        .Font.Bold = True
        .Font.Size = 14
        .HorizontalAlignment = -4152  ' xlRight
        .VerticalAlignment = -4108    ' xlCenter
        .Interior.Color = RGB(46, 82, 152)
        .Font.Color = vbWhite
    End With


    ws.Columns(1).ColumnWidth = 5
    ws.Columns(2).ColumnWidth = 60

    If viewType = "jour" Then
        ws.Range(ws.Columns(3), ws.Columns(lastCol)).ColumnWidth = 3
    Else
        ws.Range(ws.Columns(3), ws.Columns(lastCol)).ColumnWidth = 5
    End If

    With ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))
        .Borders.LineStyle = 1
        .Borders.Weight = 1
        .Borders.Color = RGB(0, 0, 0)
    End With

    ws.Cells(5, 3).Select
    ws.Application.ActiveWindow.FreezePanes = True
    ws.Rows.AutoFit
    ws.Rows(1).RowHeight = 50
End Sub

' ============================================
' INJECTION UNIT�S DE POINTE DANS FEUILLE "JOURS"
' ============================================
Sub Injecter_UnitesDePointe_Dans_Planning(xlApp As Object, xlBook As Object)
    Dim wsJours As Object
    Dim projet As Project
    Dim tache As task, assign As assignment
    Dim tsData As TimeScaleValues, tsValue As TimeScaleValue
    
    ' Variables d'optimisation
    Dim dateIdx As Object, rowIdx As Object
    Dim arr As Variant, totauxArr As Variant
    Dim lastRow As Long, lastCol As Long
    Dim compteurValeurs As Long
    Dim i As Integer, ligne As Long, col As Long
    Dim debut As Date, fin As Date, dateJour As Date
    Dim tacheName As String
    Dim minDate As Date, maxDate As Date
    Dim oldScreenUpdating As Boolean, oldEnableEvents As Boolean, oldCalculation As Long
    
    ' Variables pour les index de tableaux
    Dim idxTache As Long
    Dim colIdx As Long
    Dim arrRowIdx As Long, arrColIdx As Long
    Dim dateKey As Long
    Dim currentValue As Double
    Dim somme As Double
    
    Debug.Print "Injection des unités de pointe - DÉBUT"
    Set projet = ActiveProject

    ' Vérifier l'existence de la feuille "Jours"
    On Error Resume Next
    Set wsJours = xlBook.Worksheets("Jours")
    On Error GoTo 0
    If wsJours Is Nothing Then
        MsgBox "L'onglet 'Jours' n'existe pas dans le fichier Excel. Vérifie son nom exact !", 16
        Exit Sub
    End If

    ' Désactiver les mises à jour Excel pour optimiser les performances
    oldScreenUpdating = xlApp.ScreenUpdating
    oldEnableEvents = xlApp.EnableEvents
    oldCalculation = xlApp.Calculation
    xlApp.ScreenUpdating = False
    xlApp.EnableEvents = False
    xlApp.Calculation = -4135 ' xlCalculationManual
    
    ' Déterminer les dimensions de la feuille
    lastRow = wsJours.Cells(wsJours.Rows.Count, 1).End(-4162).Row
    lastCol = wsJours.Cells(4, wsJours.Columns.Count).End(-4159).Column
    
    ' Construire le dictionnaire des dates (ligne 100, colonnes 3 à lastCol)
    Set dateIdx = CreateObject("Scripting.Dictionary")
    minDate = DateSerial(9999, 12, 31) ' Date très élevée
    maxDate = DateSerial(1900, 1, 1)   ' Date très faible
    
    For col = 3 To lastCol
        If IsDate(wsJours.Cells(100, col).Value) Then
            Dim currentDate As Date
            currentDate = DateValue(wsJours.Cells(100, col).Value)
            dateIdx(CLng(currentDate)) = col
            If currentDate < minDate Then minDate = currentDate
            If currentDate > maxDate Then maxDate = currentDate
        End If
    Next col
    
    ' Vérifier qu'il y a des dates
    If dateIdx.Count = 0 Then
        xlApp.ScreenUpdating = oldScreenUpdating
        xlApp.EnableEvents = oldEnableEvents
        xlApp.Calculation = oldCalculation
        Exit Sub
    End If
    
    ' Construire le dictionnaire des tâches (colonne 2, lignes 5 à lastRow)
    Set rowIdx = CreateObject("Scripting.Dictionary")
    For ligne = 5 To lastRow
        tacheName = CStr(wsJours.Cells(ligne, 2).Value)
        If tacheName <> "" And Not rowIdx.Exists(tacheName) Then
            rowIdx(tacheName) = ligne ' Première occurrence seulement
        End If
    Next ligne
    
    ' Vérifier qu'il y a des tâches
    If rowIdx.Count = 0 Then
        xlApp.ScreenUpdating = oldScreenUpdating
        xlApp.EnableEvents = oldEnableEvents
        xlApp.Calculation = oldCalculation
        Exit Sub
    End If
    
    ' Lire la plage de données en mémoire (lignes 5 à lastRow, colonnes 3 à lastCol)
    If lastRow >= 5 And lastCol >= 3 Then
        arr = wsJours.Range(wsJours.Cells(5, 3), wsJours.Cells(lastRow, lastCol)).Value
        ' Vérifier que arr est un tableau
        If Not IsArray(arr) Then
            xlApp.ScreenUpdating = oldScreenUpdating
            xlApp.EnableEvents = oldEnableEvents
            xlApp.Calculation = -4105 ' xlCalculationAutomatic
            Exit Sub
        End If
    Else
        ' Pas de données à traiter
        xlApp.ScreenUpdating = oldScreenUpdating
        xlApp.EnableEvents = oldEnableEvents
        xlApp.Calculation = -4105 ' xlCalculationAutomatic
        Exit Sub
    End If
    
    ' Initialiser le tableau des totaux
    If lastCol > 2 Then
        ReDim totauxArr(1 To 1, 1 To lastCol - 2)
    Else
        xlApp.ScreenUpdating = oldScreenUpdating
        xlApp.EnableEvents = oldEnableEvents
        xlApp.Calculation = -4105 ' xlCalculationAutomatic
        Exit Sub
    End If
    
    compteurValeurs = 0
    
    ' Traiter chaque tâche avec assignements
    For Each tache In projet.Tasks
        If Not tache Is Nothing And Not tache.Summary Then
            tacheName = tache.Name
            
            ' Vérifier si la tâche existe dans notre dictionnaire
            If rowIdx.Exists(tacheName) Then
                idxTache = rowIdx(tacheName)
                
                ' Traiter chaque assignement de type Work
                For Each assign In tache.Assignments
                    If Not assign.Resource Is Nothing Then
                        If assign.Resource.Type = pjResourceTypeWork Then
                            debut = assign.Start
                            fin = assign.Finish
                            
                            ' Clipper à l'horizon visible
                            If debut < minDate Then debut = minDate
                            If fin > maxDate Then fin = maxDate
                            
                            If debut <= fin Then
                                Set tsData = assign.TimeScaleData(debut, fin, pjAssignmentTimescaledPeakUnits, pjTimescaleDays)
                                
                                For i = 1 To tsData.Count
                                    Set tsValue = tsData(i)
                                    dateJour = tsValue.startDate
                                    
                                    If IsNumeric(tsValue.Value) And tsValue.Value > 0 Then
                                        dateKey = CLng(DateValue(dateJour))
                                        
                                        If dateIdx.Exists(dateKey) Then
                                            colIdx = dateIdx(dateKey)
                                            arrRowIdx = idxTache - 4 ' Conversion ligne sheet -> index tableau
                                            arrColIdx = colIdx - 2   ' Conversion colonne sheet -> index tableau
                                            
                                            If arrRowIdx >= 1 And arrRowIdx <= UBound(arr, 1) And _
                                               arrColIdx >= 1 And arrColIdx <= UBound(arr, 2) Then
                                                
                                                If IsNumeric(arr(arrRowIdx, arrColIdx)) Then
                                                    currentValue = CDbl(arr(arrRowIdx, arrColIdx))
                                                Else
                                                    currentValue = 0
                                                End If
                                                
                                                arr(arrRowIdx, arrColIdx) = Round(currentValue + CDbl(tsValue.Value), 2)
                                                compteurValeurs = compteurValeurs + 1
                                            End If
                                        End If
                                    End If
                                Next i
                                
                                ' Libérer tsData
                                Set tsData = Nothing
                            End If
                        End If
                    End If
                Next assign
            End If
        End If
    Next tache
    
    ' Réécrire le tableau en une seule opération
    wsJours.Range(wsJours.Cells(5, 3), wsJours.Cells(lastRow, lastCol)).Value = arr
    
    ' Calculer les totaux par colonne en RAM
    For col = 3 To lastCol
        somme = 0
        arrColIdx = col - 2
        
        For ligne = 5 To lastRow
            arrRowIdx = ligne - 4
            If arrRowIdx >= 1 And arrRowIdx <= UBound(arr, 1) And _
               arrColIdx >= 1 And arrColIdx <= UBound(arr, 2) Then
                If IsNumeric(arr(arrRowIdx, arrColIdx)) Then
                    somme = somme + CDbl(arr(arrRowIdx, arrColIdx))
                End If
            End If
        Next ligne
        
        If somme > 0 Then
            totauxArr(1, arrColIdx) = somme
        Else
            totauxArr(1, arrColIdx) = "" ' Laisser vide si total = 0
        End If
    Next col
    
    ' Écrire les totaux en une seule opération (ligne 3)
    wsJours.Range(wsJours.Cells(3, 3), wsJours.Cells(3, lastCol)).Value = totauxArr
    
    ' Appliquer le formatage de la ligne 3
    With wsJours.Range(wsJours.Cells(3, 3), wsJours.Cells(3, lastCol))
        .Font.Bold = True
        .Interior.Color = RGB(255, 230, 153)
    End With
    
    ' Restaurer les paramètres Excel
    xlApp.ScreenUpdating = oldScreenUpdating
    xlApp.EnableEvents = oldEnableEvents
    xlApp.Calculation = -4105 ' xlCalculationAutomatic
    
    Debug.Print "Injection des unités de pointe - FIN (" & compteurValeurs & " valeurs injectées)"
End Sub

Sub InsererLogoOmexom(ws As Object)
    Dim base64Image As String
    Dim byteData() As Byte
    Dim xml As Object, node As Object, stream As Object
    Dim tempFile As String

    ' === Logo Omexom en base64 ===
    base64Image = GetBase64()

    ' Conversion Base64 ? octets
    Set xml = CreateObject("MSXML2.DOMDocument.6.0")
    Set node = xml.createElement("b64")
    node.DataType = "bin.base64"
    node.Text = base64Image
    byteData = node.nodeTypedValue

    ' Fichier temporaire
    tempFile = Environ$("TEMP") & "\omexom_logo.png"
    Set stream = CreateObject("ADODB.Stream")
    With stream
        .Type = 1
        .Open
        .Write byteData
        .SaveToFile tempFile, 2
        .Close
    End With

    ' Insertion du logo
    ws.Shapes.AddPicture tempFile, _
        LinkToFile:=False, _
        SaveWithDocument:=True, _
        Left:=10, Top:=5, Width:=120, Height:=40

    ' (Optionnel) suppression du fichier
    On Error Resume Next: Kill tempFile
End Sub

Function GetBase64() As String
    Dim parts(93) As String
    parts(0) = "/9j/4AAQSkZJRgABAQEA3ADcAAD/2wBDAAMCAgMCAgMDAwMEAwMEBQgFBQQEBQoHBwYIDAoMDAsKCwsNDhIQDQ4RDgsLEBYQERMUFRUVDA8XGBYUGBIUFRT/2wBDAQMEBAUEBQkFBQkUDQsNFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQU"
    parts(1) = "FBQUFBQUFBT/wAARCADbAqgDASIAAhEBAxEB/8QAHwAAAQUBAQEBAQEAAAAAAAAAAAECAwQFBgcICQoL/8QAtRAAAgEDAwIEAwUFBAQAAAF9AQIDAAQRBRIhMUEGE1FhByJxFDKBkaEII0KxwRVS0fAkM2JyggkKFhcYGRolJicoKSo0NTY3ODk6Q0RFRkdISUpTVFVW"
    parts(2) = "V1hZWmNkZWZnaGlqc3R1dnd4eXqDhIWGh4iJipKTlJWWl5iZmqKjpKWmp6ipqrKztLW2t7i5usLDxMXGx8jJytLT1NXW19jZ2uHi4+Tl5ufo6erx8vP09fb3+Pn6/8QAHwEAAwEBAQEBAQEBAQAAAAAAAAECAwQFBgcICQoL/8QAtREAAgECBAQDBAcFBAQAAQJ3AAEC"
    parts(3) = "AxEEBSExBhJBUQdhcRMiMoEIFEKRobHBCSMzUvAVYnLRChYkNOEl8RcYGRomJygpKjU2Nzg5OkNERUZHSElKU1RVVldYWVpjZGVmZ2hpanN0dXZ3eHl6goOEhYaHiImKkpOUlZaXmJmaoqOkpaanqKmqsrO0tba3uLm6wsPExcbHyMnK0tPU1dbX2Nna4uPk5ebn6Onq"
    parts(4) = "8vP09fb3+Pn6/9oADAMBAAIRAxEAPwD9U6KKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACi"
    parts(5) = "iigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAoooo"
    parts(6) = "AKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACi"
    parts(7) = "iigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAoooo"
    parts(8) = "AKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKTNFAC0UUUgCiiimAmRRTTSbqBj6C2KbuzTW+XrUNsT0RJuo3DGaj8"
    parts(9) = "wClUhh7VoJX3H0bhTd3btRuHep9B3Q7NG6omkA/iAFJ9oj6b1/OqsyOePcm3UbhUHmrn76/nTw6n+IGizEpokpN1IxLL8ppPrU37mo4NS5FRhhnFG8KeaYrolopm8LSF8c9jQK6Hht3SlqPzFXihm6GlqMfS03cNuabuLcjpRcB+aNwpi/MCTQrBmxT9RD91G4VE2V5P"
    parts(10) = "SlVty/LT03KJN1BYCo0560kkiocGpfkSvMl3CgnFMU5XNC5K/NzTGO3flRuFRNMkfBpVkVulFmJtdCXdRuqPcAcUFxuC0/UNSTNG4Uxm28UMwXGah3C/QeTjrQWC4zTdwYZ7UwfLkt+FNbXGTUlRC4VulO3Hr2qhXQ7dS1HuGKTzAtQ2DaRKWpAwpnmD8aFI5NCY1Z7E"
    parts(11) = "m6k3VGrbj7UB/mIq7D2JaM0w549KRnC1GtyW7EmRRkVGJB3pNx8zjpTjdg2iXNLUW/nFOZti0wWo6gMDTQ4YU3zVXjvSWo9iTdRmojKo/iH50nnKR99R+NXymftET0mahWVe7qfxpwkVujA/jSsx8yZJSFqYJAGxmmmQR8tz6UnoUmmS7qWo1kVl3UplWldBdDg3OKXN"
    parts(12) = "Rgg/NR5gourjZLSZpu7uOlIzhcE0a3Akpu70ppbzF+WkU469aetxktJnmm7qN4pK97Etpbj6TcKazjFRNOq9aXvXKtpcn3Um6o3fCg9qVWEi5FXbS5Cd3Zj91LuFRqwancUrorToO3UVE0oXg1IDxQJbC0hNL96kHrQMNpopc0UrgLRRRTAKKKKAGN3qNutSNUTUCGOx"
    parts(13) = "XBpDN5h6U9lDdTVW6u47dS0h2Rjq1Hxe6iKtSEY3kWTwvtVa51azsYyZ7mOEd97AV498WP2grPwVaMunSR3lzggopyRXyz4q+NOu/EyOW1m83TWckAqSK9zDZXWrq70R8hjeJsLhbwT1Psbxl8dND8Kwllu4blh/Cjg14z4i/bs0zS1dU053YdCDXzZpnwX8TXl0bgXl"
    parts(14) = "1eRscgMSRXf6D+znf6tIjXdqyjvuWvoaeX4GjZV9z5WWdY7FS/2fYv6p+3z9scrFp80fPWsW5/bUu7hgyRyoPrXoVr+yDa364dBGfXbUeofsU2hjIST8hXrxeTxtFkTrZq1c4O3/AG5pbVsSJI+Peuh0f/goHawzIJbN3Fcz4o/Yvk0+3eW33Sv6Ba8O8RfBzxD4XmkZ"
    parts(15) = "tKk8tD94pXXHC5ViHyxMlmGYUdZn3d4T/bY0TxFIiNbfZy39417R4d+JGja9EskeoQbm52bxmvyB864hkKCRreYfwjjFaXhnxlrfhe8+1pqdxIIzuEe81hieFaUoc2HO3D8TTUuWsz9l4biOePfG4f8A3af5o6Ec18LfAz9sqW4uoLLVlEMX3TJIa+0fDviKx8UWMd5Z"
    parts(16) = "zrMGG75TmvzzH5bWwMrSWh+gYHMaGLjo9TZOFj5OTUF1fR2kBeQhVAySaay7280nkfw15b+0V4wk8K+DZmj+VpY2AIPTiuTD0vbTjDudeKqKhTc0UviD+05ofgW6NuzJcsP7rVxum/tt6HeXcUJttm9tuSa/PHWNXv8AVtSurme8klPmHCsxPeqxW4bZOszIyHcMV+n4"
    parts(17) = "fhWnUo87R+a1uJKkK3LfQ/Zbwr4qtvFtlHdWzqUYZwDW+w29OlfDP7GvxsklRNFv5MSO21d55r7iSRZIxzketfneZYH6nXdPofoGX42OLpKVx27ecCkbjgdaFxGCc5prsAvmZryZR5kexHzFyY/vc5qOe6itUMkjCNB3Y4FMmuhHC0svyIo6mvk39qT9piLwvYz6Vpkq"
    parts(18) = "yT9Moea9HBYGpjJqnBHj5hmFPBxvJ6nrnxC/aM0TwLMEM0Vy2cEKw4rgJP23tDVtptcn61+d2sapqmtXz3899NJ5x3bGYnFRRmU/MZmzX6lhOEqM4J1EfnWI4krKXuPQ/RfT/wBt7RL3VLeyW02tM4QHNfRGi6x/bVjHcoNquu7Ffj58O9Pn1/x5o6B2AjuFz+dfr34Y"
    parts(19) = "tV0nw7ap12wrn8q+LzvLqOAqKFNH12UZjPGQ5ps4H4ofHfT/AIaAG6QSHOMZrzcftyaGyn/Q8Y96+e/2zPE0uoeIJLeKQjZJ0Br5yWSSSMfvCDX02WcN0MXRU5o+fzDPp4ao4QZ+iUf7ceiZ5tOPrTm/bo0EDP2UfnX586bYz6nMlvHIxdjjivU9O/Zn8QajbxzLFOVc"
    parts(20) = "Z6GrxOR5bhZWqaGOEzrHYj4WfVf/AA3RoUjf8emP+BVKf259CVc/Zen+1Xy5L+y14hRBtgnJ/wB00z/hlnxFKOYJx/wE1xfUsmtv+J2rF5lds+px+3PoUkfy2uD/AL1RR/t26EVcm0+515r5cP7LniO34WCdvwNTx/sra9N8pimXf14NZPA5Ta6f4m9PGZjJ6n3R8KP2"
    parts(21) = "gNP+KkiraW/lhu+a6vx547/4QPTpbtrdrlFXdtWvLP2Z/gcfhzoMM9xI32gfwsOa9u1jR7fXLGe2uI1dZBt+YZr42v7CFflh8J9RRlVlTvLc+bv+G6NIXeZdPaJlbGxjzTm/bi0QhW+ycHtmvD/2pv2e7rwtfT61pcbSQjnai8V8wQ+fIHWVzHKBynpX6Bl2SYDGwU4n"
    parts(22) = "xePzbFYabUnofp/4H/ai0PxlfRwqyW27+81ez2t5FqEKywSrIh5BU5Br8XLC8v8AS3F1DeSRMh4CkivuL9lX9pCPUoItH1a4EZjUKryNyxrgzjhv6vD2lBaI6Ms4g9rPlmz7Izu4HFDZRfeoLW8jvoUljbKsMgjvWV4l8X6Z4YtJJr26SHaM4Y1+fwpVJS5EtT7ieKpq"
    parts(23) = "HtJM1W1COziZriQRr/eY4ArxL4l/tV6H8Prx4MLespx+7bNeB/H39r6e8afRdKXdDLlDPGelfI+papdXl+80t491JIfusSetfe5dw3KcfaYhaHx2Oz6N+Si9T72/4b90Taf+JeQe3Nek/CX9pW3+J16tvBp7xq3/AC07V8EfCH4A678SdWgF5aS2to7D95t4xX6IfCz4"
    parts(24) = "O6X8I9GjiR1k8sZMpHNcea4TA4X93R+I6suxWKre9U2PUY5M8kdar3mrWljk3FxHEP8AbOK8h+LH7SWieBdNm+xXcN1dop/dgjOa+H/iT+15rnxCSaySN7GRiQNhINebg8mr4pc1rI9DGZtSw65YvU/QPxV8btC8NqxF5DOV6hXFeNeKv26tE0bfHHa+Yw7g18L6T4b8"
    parts(25) = "d6/M0kP2y8ST3Jrv/Cv7M/ibxBIrX9pPHu67lNfSRynL8KksQ9T5x5ljcQ/3ex69qX/BQC2lZvLtmXPvWK/7c0koLKHX8aXS/wBhVrhjJK7Ju7EVcvv2DBbqWSZmPpiu6EcoT1OabzPdFO1/bqkjb5o5HH1roNG/b6ghkAltJGry3xN+ynq+g7vstlJcY9Fry7X/AIZ+"
    parts(26) = "I9BkLzaTJHEvVildccFldd2gcf1zMaOsj728IftraN4kkWF7b7O395jivZ/DvxT0TxHGpS+gDH+HeK/IcXUiL5SzG3mH93g1p+HvE2seGrpbtdUuG2HITeeazxHCtKpDmw5pR4jqQny1WfsnHcR3UY8lgyn+JelLzCwUnO6vgv4K/trXhuodG1OHy7dcD7Q5r7a8I+LN"
    parts(27) = "P8XabHc2NwtxkZbac4r88x+VVsDK1VaH3eBzGji1o9Te3bH2dqhvL6KxiaSVgqjkk9qdKwVeDlq8e/aU8ZP4V8G3Ajba8kTAEHkVw4Oh9arqjHqdeLrfV4ObG+Pv2ndD8DzNGWjuSvZWrkNL/ba0TU76G3+y7PMbbuJr869Q1S/17UJ7ie9kf5z8rNnvTjNLAFeOUh05"
    parts(28) = "GK/TafCtOVO7Wp+e1uJJxq8qeh+yPh3xRbeJNOW4tXV1YZ+U5rZztQE9a+KP2Mfi+8lkmj3sxaZm43nmvtQsJo1YHg1+fZjg5YGq4SR9xl2OjjqXOnqPDg0uKYke2n8DvXjqV9Ue1KKaQ2T7uaqzXEcMbPMwRFGctxVi4uI7eFpJGCxqMlj0r5C/am/acHhi0m0zSnEr"
    parts(29) = "/c3Iea9bA4KrjaipwR5eYZhDB0m2eu+OP2lNC8GzNEZI7krwQrVwR/bk0PzCgtMe+a/PDUda1HUL6W6nvZZDcNv2sx4zSN5rw48wg9c1+lUOFKLj761PzmrxNNP3Wfo7of7aGja3rFrYx2u1pm25zX0Rpl8NTtY7lOFYZxX5C/Bq3uNU8faWu5sJMoz+Nfrb4TtWs9Gt"
    parts(30) = "kJ/gH8q+MzzLKWAkowPrsozCeMjeRtYDc46Uscnmc4xSM5XGBkUnevkUlI+ujHQlBFKT2qNakxzVg1YRaKdRU2EFFFFMApKWigCKSmY4zT5DTQaBWILiRYY2djhV5NfOXxy+NEmlyS6Zp0gZiOxr2z4gar/Zfh68kBwRGcV8D+JtWbWtdluWYsdxFfU5NgY1588kflHG"
    parts(31) = "WcSwdH2dJ2ZnzW8uu6os0bvLdytkoTnmvePhr8BH17yrrWYDC3UYHasj9nnwHb6xq8t1cx79h3Lmvr/T7WOG3VUGMDFejmWO+rP2VLQ+a4ZyaeZf7RidUYPh/wAB2Ph+1SGGNSqjHI5roYrGNFwI1H4VP81SKDjmvkKtecvem7n7JQy7D4eNqUbES24/uj8qc0KBeVFT"
    parts(32) = "I43Ypkqkniub2jep6EcPFLUqyWcJ5ZVI+lYfiLwPpniizkt7iFCrDBO0ZrpyqeWN1MVV2ny60jiJ0mpRdjCeBoVFyyR+fP7Sn7Lsvhtp9X0KBpDnPA4r5WVWhkkhnG2ZOHX0NfsZ4002LVNBuYZ1DL5bHke1fk58UdMh07xxqaW42r5zZ/Ov2PhfNJ4qPs6vQ/J+IMBS"
    parts(33) = "w0rwRyaxq0wBkaNQcgqcV9T/ALI37QF3ourHRtWm227MI4Sx618v7Y2fFaHhSSe08caGyMQPPXGPrX0WeZfSxGHc7HkZPjZ06yimfsrbzR3dqkqHO5dwr5a/bd8VCw8O2lurYLEg19BeAbqS68NW0rnnyh1+lfDv7bvipb++jtFfJik5FfkeTYP2mO5P5T9OzTFP6om+"
    parts(34) = "p8sTMFuH54Y5rVsdJv8AXoWXTovNMI3PgdqyQolYSH7vevov9jnwafEXiLVYpo90MkRAyK/aMVi44HDcz2R+SUMO8XXsjxXwP4wuPB/i221AMYxbOA/YZzX6jfBP4kW/xA8I2lysoeVlGcGvz/8A2mvg0/gHxBtghKWsxLtge9dB+yX8Y5PCfiJNOupitj9xFJr4LNcJ"
    parts(35) = "DMsP7elq9z7bL8TLA1fZTP0rAyMUlwyxw/McBeTUGn38V1p8NyrDbIgYHPrXzx+0l+0RaeBtNns7SfF7yp2mvzfDYSpiavs4I+9xWOjh6PtLmd+05+0jF4R02XTtJnV7llKkA85r4Lt7XWvit4ilmdXlvWy2zqMUmoXmtfFDxC43PNcXD/uj16mvrr4TfBlfhV4EPiDX"
    parts(36) = "YlF+UILH6V+mYajRyqnFfaf4n5xi6tTMk5dD421Cyn0m8ktLhdjxnawquSinIPFaPjjVP7V8Z6pIpzG8pK/nWT5flAK1fo9Cb9jc+KrR/ecp7n+xx4NHjDxnPK6ZFs4YfnX6WXjrYaGSxwqR4/SvjL/gn/4d+x3mq3TLgSLkGvrH4oagdN8IXcgOMKa/Es8qOvmPsuzP"
    parts(37) = "1jJqaoYLn8j8yv2htWbUPiLqiZ3IspxXmY4YbeldF8R71tQ8dalITnc5/nWAsfyn1r9hy2n7HDxSPy3HzU8TK/c7X4FaZJrnjuOELuVJAcfjX6veFdLitdFs08mMbYlB4HpX5M/CnxfF4D8QPfSnB619L2n7cFtbxxwGdsKMV8FxHl9fFVl7N6H2WR4yhh178T7n+yw/"
    parts(38) = "88k/IU/7LD/zyjx9BXwvJ+3HCsxAnbFPX9uKHb81w2K+N/sHFd0fYf2xhov4T7jazh6iJCfoKT7HH1EMefoK+HF/bmt1b/XtitHwr+2kuueJLLTlmdjcOFWsauT4mjG8mjalmmHqSson2gzDdswE9hT1XzAwPQVn6LcG8sI55OXZQavbXbOzgd6+drRcX6H0VPllHmRj"
    parts(39) = "eJfDNn4s0mayu0VoXBySK/Oj9pH4B3/gPV7jUdPtmNpIxO7HGK/S5WX/AFQ61znjrwbp/jPRLiy1CJZAyFVz2OK+gyjNKmBqJ393qfPZpl9LFU20tT8dUuDdBiPuKcN9at6RqUmi6hDdW0jK8bBvlOK9F+Onwrb4Y+LJLaPaLN2LlVNeauImXMY4PSv3XD1oY+gnbRn4"
    parts(40) = "rWjLBV3yn3B8Kv2tbWz8Dy/2vcrHewrtjUnrxXzx8aP2hNc8fajPHG7CxJ+VlNeQIm7Kyk4PvXQeDfBt94+1WPS9KIDKw3bvSvHlk+FwPNiOXU9uOa1MRFUUzF0exvNduDa2Yae6mPAIzzX1f8B/2QBri2954iheFxhsEV7D8Cf2WdK8MW9veapbK18AGVsd6+kbXT0s"
    parts(41) = "FEUKhVXpgV8TmnETlH2OHdj63Lcl5kq1TUy/DvhWx8J6XDaW0KJFEu3dgZ4r5w/ah/aMXwfaz6RpcytedCuea95+LXiR/DPgzUbtG2yRxlga/J/4h+Kbrxn4tuNRuZDJuYjrXBkeBeOq+2q62OrNsSsJD2dLQo6peap4x1gzRySy3k7ZEeSRk19N/AH9kefxV5OpeJrd"
    parts(42) = "oJOuMY4rnv2PfhlB4u1uS7u4xILdwy5r9G9MsU02zjjgUKqqBwK9TPMyWC/cYfRnn5RgZYyXtK2qOY8HfCXRfBdskVnAj7Rj5lFdgtrDEvEKD6KKkhzyT1peerc1+czxNTEPmnK7Pvo4alRajBWAxxhRhF/KnmNGXlFpFTd9KXa2eelYttI7vZxSIJbeFlwYkP1ArlvF"
    parts(43) = "Hw30jxdYywXcEaq4wdqiuvdkXr1qPjaSK1p16lF80XY5KmGpVlZo/PX9oz9lMeFPN1TRImkUnPTtXy9NG6StBMNrxnBFfsT440q31rw7dwzKGURseR7V+T3xT0tNI8ZXscQwpmbp9a/YeF82niI+zqPVH5NxBl0KEuaBykkY2gBijDkFetfSH7Jfx11DwnrMej3Mube5"
    parts(44) = "cJ85zxXzeylpiK2PAs00XjXS2iONso/nX0mcYOjiaMuZHh5XjKlGrFJn7G2M0V9DHcRNuRlBr5Z/bm1xbPRbWEPhpFIIr3r4V30t14Pt3kPzbO/0r4s/ba8Um81SC2Z92xyBzX5DkmF/4Ufd6H6dmeJ58F6nyraqVkf+8WJra0fw5e68s7WUZkMQy/tWZGPLj8w9c19O"
    parts(45) = "fsf+DoteuNUSePcJkwMj1r9ox2JWBoe0kflOGw/1qq4o8B8D+Nb/AOH3imC+jJRomwR2r9R/gj8R4fH3hKzn80NcsgLgGvz8/ac+EkngfxeBDFstGG84HFdd+yH8a08I+ITYajMRbSYjjBNfn+b4eGZ4d4ilrY+0yytLL6yoy6n6O/NsJxzUU8hjhLnjaMmo7a/S4so5"
    parts(46) = "kcbJFDA59RXzl+0h+0hbeCdNn0+wn23uCjbTX5zhMDUxNVU6aPvMVjoUaXPcyf2mv2lofC1lNpWk3CvLIpVsHnNfCM0ms/EzVpWgV57nJZl68Ut9fan8TPExgbzJru5f923Xqa+wfh38ErT4X/D0azqMAXUpIiGcj2r9SwtOjlKjTXxP8z83xNSrmV29j4uutLls5DHM"
    parts(47) = "u2SLhh6GmNJmL5av+JdTM3ifUjnMRmJ/WsuWQbtycRniv0CnLmo83U+GrUuWtyHtv7H3h7/hJPGAmZc/Z5QenvX6h20YhtY1HGFAr4X/AOCfPhQKdXu5o++VJr7rRdwHpX4RxFXdbFOLex+08OUFCjcljJ28ik2hqA3BFIo/OvlOXQ+xcrPQkwBQOtLjim0w1uKOtFJR"
    parts(48) = "SGPooooAKKKKkCGWmLTpKatUOxwXxejZ/C18VGcRn+VfBMuVmmOMN5h4/Gv0b8VaUNW0m5tyM+YhWvgr4oeE77wb4skU2zfZMkmQjivtcirRUuVs/CuNstrVpe1iro9f/Zr8RW8NxLC5CMRjmvqaGZWjBXoRX50+G/EreH71Luyl8xs7mRTX1P8ADH452+tW8cWoMtq+"
    parts(49) = "MfOcUZxl05y9pDU04TzuGGh9Xq6Huqtu7VMPu1iaf4s027AEN1HIf9k1o/2hG44bivkJ05rRo/YaeMoTV1JE6x85p0lV1uk/vU37VFGCWeo5ZbWN5V6TXxE8nzJTRII1NZl54m02yUtPdxxgf3mrxz4sftLaP4JtZGs7iG7kUfdVs114fA1sRPljE8rE5hRw8XLmOi+O"
    parts(50) = "PxU0/wAAeF7ieaRGkYFfLzzX5b+LPEH/AAkvia+vVBVJpCw/Ouu+Lnxg1H4t6xLcSyPb27HiLPFeeTXEcAROAema/Zciyl4Cnzz3Z+S5tmTx0+WKHzcdOtdz8EfBN94+8bac0MT7LWZSxxx1rnfB/gzVvGOtRW1paSTQsR86jIr9IP2b/gfa/DfR0upYVa5uEBO5eVNV"
    parts(51) = "n2cQo4d0ou7OjJcqnUqqo0es6fZrofh2OPO0JDg/lX5e/tH64dU+IOpwsdyrKcfnX6a/Ey+/s3wndShtoCnmvyb+Jl6dS8eam+d37wnP418rwrTdWvKqz6XiCao0VTOYuI2jszhsciv0A/Yj8IfYdIj1FkwZo/vYr4FW1a+njgB+8RX6s/s16KmkfCnRdq4cxDJr2eKs"
    parts(52) = "U6VH2Xc8Thuh7StzmZ+0n8Obfxf4PvJRbiW6VCEbGTX5j6lpd94N8SKgZoJYJdx7Hg1+zs1tHMhSRQ6t1Vua+Af2zPgr/YuoT+JLRMJO/wDq1HAr5vh7MEpfVqj0PoM8y6Sft4Gn4f8A2ult/hvPZOzfbI4tiPnnpXyr4j8Qax8SPERlmMt2JXwBye9YkK3GouLWzBlm"
    parts(53) = "PHlr1Jr7Z/ZN/Z3C2tvrOsWux2GfLlXpX1eJjhMrhKpFK71PBws6+PkqctkbH7LP7NaaNZxatqsKyOQHjDDkV2H7Z2vL4d+Esq27eWdwAUV9E2NnHp8KQQRBI1GPlFfFv7dms+dp82nb/lznbXwuDrTzPHpyeiPp8woxy/B6I+KLWYXm65YfM3OTTLyVmiRu+8CltVAh"
    parts(54) = "jXpgVas7U6jqMNoF6uP51+2T/d4W/ZH5dCSr1z9Hf2PvD403wjb3QTBmiBziuu/aY1r+yPh1fSbsECtL9n/SxpPw90hduP3Sj9K8j/bY8Umx8LXVkvRlr8Si/reaXfc/VlD2OX2R8BaleG81me4PO85qvd3QgTzAMgVFbEzxBz3p0luXwj8Akfzr9ti3CgrI/I5pPE6s"
    parts(55) = "9x+D/wCy7qXxXslvo5xDDIu4bq9L/wCHe9+0bf6XEW7HNfQP7Kum2un/AAy0uSNxvaPnFe5RTIOM5r8dzTOcT9akqbskfqGW5bh6lJSk9T4Oj/4J66h5Y3XkRb60v/DvbUWbP2yLH1r7z81FOd2fal+0R4+9XiPNsXe9z2/7Lw3c+DJf+CfF9t4u4vzrsvhT+xCfBOu2"
    parts(56) = "+o38kVw0Thl9q+vvtCMdu78acCNpw26s6uZ4mpHlkzWlgaEHzRYyzt0t40iUYVVA/KpHky2E4A61GwYLx1rE8UeNNL8I6dJPe3UcDhchXOM15Mac60rLVs9j20KUNXaxr3l5bWMLSyyLFgcsxxXzR8ev2sNN8D281nbHzrh8oHjOcGvHPj7+1vc6tNPo+mZEbZHnRmvl"
    parts(57) = "PUL+71e8eSe4a8mkPyoxyc1+g5Vw/tUxGh8Bmmdc16VI6Hxp4y1Tx5qj6lfXbTJnhGPaufjjuJMn7LIkQ5DkcGvYfgr+zfq/xEnie+glsoS38QIGK+vdZ/ZX0qTwOmmRIi3EMZ/eBeWOK+uqZ3hsuaoI+a/suriYe1aPzbDeYT3wa6PwT4mn8B61BqcEhXc43BfSrXxM"
    parts(58) = "+Hl98NddmtZoH8lnJ3sO1crbydZD86HpX0nPDHUfde585KlPCVbyWx+pPwL+Nun/ABC0GFRIqXCIFO48k4r2SGM7QS2fevyF+F3xHvvAPiK3uIpn8hXyyg8V+mXwb+L2nfELQLd/tCC6KjMYbmvxrPMmng5upBaM/XcjzSFemqcmU/2krSab4d6r5QJ/cngV+UTRutzK"
    parts(59) = "kgKMHJwfrX7P+JtKi8RaPcWMqgpKu3kV+aP7SnwTv/Bfia4u7K1d7XOcqvFenwzjYUb0p6M8viLCzm/aQ1R3/wCw/wCNrTS9Ru7SdlRpGwNxr9AIZUnhVozkEZ4r8ZPDviWXwfqdtfWs5WaJgzxKe/pX3B8B/wBryHxB5Fnq220XAUvIcVfEGUTrv6xS1uRkWZxor2U9"
    parts(60) = "D7AQU/aHU1z2m/EDw9qUatb6pbyFuyuK2odQt7hd0MiuP9k1+c+xnT0aP0CNalLVSJFyTjOMVKrfLiq7Tj6VG+qW1mhMsqoP9o1l7KcmaLEU72bJmYbiCM1FN8jh8/IOorF1H4heHNOR3m1W3RlHQuK8A+LX7X1h4VhmTTvLvCAQCpzXo4XLa+Jlyxiefiswo0FfmPQf"
    parts(61) = "j58VLDwR4VmlFxGZpAU2BhnpX5g+J9bbXtavLtyW3yFlJrZ+JHxO1T4o6vNezXEkUMhyIdxwK5MNHDiN2AJ7mv2fJck+o01Ob1PyXOcyeLm4wQzadu/PNd98FvCtx408daW9rEwihlHmYHHWuZ8I+C9X8VeIEtbSzkngYgb1BIr9F/2bP2d7X4ZaULqcLLcXIDncOUNR"
    parts(62) = "nWbU8JScE7tlZNlc601OSPW9Psk8N+FxGo27Iv6V+Yv7Q+vP4j8bX8ZbIhmPX61+mHxQ1RdD8J3EucYUj9K/J/x5qJ1Txpqko6GUn9a+Y4XoOtVlX8z6PiCoqFD2aOaupGWJQP7wFfoL+xX4Z+yaOLxk/wBZGDnFfADRtcXEUaLuJdePxr9Wf2bfDaaP8N9ImxteWEEj"
    parts(63) = "8K9biyu6dJQvueNwxQ9rU5mYP7T3wpTxz4Nu5oIwbxVwrY5r80LrSb7wn4njh3NDLZSbmJ4zg1+0E1uk0DRyqHQg5BHFfnj+2Z8JX8O6k2t2cJCTyZbaO1fOZBjeZ/Vaj0Pos9wMqbVeHQ6bR/2wEh+Gs0LFhdwx7FYnngYr5T8SeJ9T+Ifib7S4kuxcybQoycZNYcaz"
    parts(64) = "6xOlpZDzHbAMS9zX27+yn+zXHaQwa1qlvuaQA+VIv3a+pxCwuVKVWNtTwMN9Yx7VORu/sv8A7NtroNnBquq2qzTsBIhYcrXXftia6nhj4bKITsBbaAK+g7Gyh0+BIIkCKowAK+I/28vGBn006QOkb18Tg8RUzDMYTlsmfV4vDwwOEaW7PjKQG7upbgnO87qbcRlhAqcZ"
    parts(65) = "kAxTrX5Y4/oK1NJsVvtUtgp3N5i/L+NftNRqlRd+x+RK9bEH6Nfsk+FB4f8ACMUwTabiNWNfQ6/IuK8/+Cuniy8B6V8u0+Qv8q76HLHniv53zKp7bETl5n71lFF08Oh0eck09fm5pPWlWvH1R7F9R1OxSbaUitCwC0Ui0UgHUUUUwCiiigCOSo6lZaZt9qAGEDqelcP8"
    parts(66) = "RPhvpvjzTZbeZFDMPvAc13flkqwqGOFVG3HzVpRqyoz5os48Vh4YqHJNHwB8QvgV4j8B3k8+h2cl1DnJJGRivNT4h1PTGL6tus5FOCBxX6j3WnxX1u8UyKysMHivL/FH7N/hLxMrm7tMu3cV9jh8+fLy1UfmuN4QhKTnSdj4w8N/HxNAZWivGcj+8a720/bEnjUKXTFd"
    parts(67) = "h4t/Yr0lt39lWpHpXm2ofsWa9uYW8BxXrRqZfiPenJI8ZZTjcK7Qub8v7Zc0anayZrmda/bQ1OZGVCo9MVmTfsS+LSx2w8fWmD9iDxazAmDitoxyqDupI2WFzCWjuefeL/2mPEGvh4yzJGe6mvMLrWptZnMs1zI7E52sSa+tND/Yn1ZdovbbI716h4T/AGLfD1uytqFp"
    parts(68) = "k9671muBwseaDTCOU4us7TufBel+Eda8QyLHp1s02fRa9y+F/wCyLq3iiaJtbtJIIyRzivuTwv8AAHwn4V2PZWgR17nFeiW9nHawiOJFUKOOK8LGcWVKkeWkrH1GB4XpwalUZ5T8K/gDo3wxs4/ssayuo6uM16nGq4wFC+wqZVf+LpSeSc5FfD1MRPEy5qjuz7ejh6eF"
    parts(69) = "jaCPIv2kdeXRvh1ekttOK/KzVtSW78QXcxOd7f1r9ffib8PofHmgy2E6bw/avCP+GKdB4c2nz96+yyHNKWXp3Z8TnGW1cdK9j4V+G9rDrXj6xsGOfMI4xX61fDLTRpPgnTrVRxGgFeMeD/2Q9A8P+JLfVo7XE0PQ19G2OnrY2iQoMBRiuXiDNI4+S5XsdeR5ZPBR1Q7b"
    parts(70) = "uG4dRXFfFL4f2nxE8Nz2FyoLBSV474ruFjdPpQIdrFu5r5GnUdNqS3R9ZWpKrHlaPiL4MfsgHQfGM+pajbsqRz7k3DgjNfZ2naZBYwrbxRrHGo/hGK0trEEMBSeWsfWurE5hWxTSn0ObD5fSw93EiklEK4Xle9fm5+2t4mS48d3NiHz7V+k0luGUgdD1r58+JH7LmleO"
    parts(71) = "/Fcmq3VvvLDrXq5LiqeExHtJs8zOcNPFUeSKPzHSaJVTJxt9q7L4PaePE3xEtrVRuUkHpX29/wAMVeH23ZtOK6X4d/spaH4J8QR6lBbbJF74r77GcR0alBwjI+KwuQThVUpI9f8Ah/pp03wvY25G0RxgV4x+1l8L5/GXhG8ls42lnxwor6Mt7YW8KxqMKOKiu9Nju42h"
    parts(72) = "lUNE3UGvy+hjJUcR7VH6DLBKdD2Z+L11ot14duXs9QjMLRcDIqtJfQSDEjbT24r9SPGn7LvhHxZdSXMtoDMxycCuNb9ivwnJJuez6V+nYbiqmqXLM/N8Rw1UdfmifHfgj9qvxL8PtPi0+zj3WsYwhIrqJP27vGkTApbAx9zivp3/AIYu8MPlWs/3Y6Ukf7GfhnBRrP5K"
    parts(73) = "8ivjsvrzc5WuezHLsTRhyxufMbft4+MdoZbcHPtTf+G8PGWRm3AH0r6e/wCGLvC+7C2fy1In7F/hXaQ9n9KxeKy22iQ/qeMa6ny8P28PFwf/AFIx24r339l74/8AjD4oXc/9p2myFW+Vsdq6OP8AYp8JCZd1n8oOa9h8C/CPRPh9brHpVv5XHJ9a8rGYrCSjamj08Fhs"
    parts(74) = "RTd5HP8Axi+Nmn/DLRpJ551S7C/cNfnp8XP2itZ+KN9LHJI0VqpOxkyMiv0V+JnwJ0T4nBv7UhMhb3rzX/hi7wpCoSO0+UcVplOKwuFlzz3Ncww9fELlR+dmiaHqHiK8W109Gup3PGRk19cfAH9jqPUpoNT8QRPA6EOFI6mvorwP+y/4V8HXiXsFoFuF6HFey29kltAs"
    parts(75) = "cahVA4wK7M04idaPs6GiODA5Co1OeoZOieG7LQbOK3toUREXblVrRIzxjIqcwkDAp0ce3O6vhalSc3zNn2UMLCmrI8R/aB+A+m/Ebw/c3CRD7aqHaFHWvzQ8aeF7rwDrtxpl9G0UUZIBIr9m2gB4AyteNfFb9mXwz8RZGuZ7UNdsclgK+wyTPp4FqMndHzObZNDFpuCP"
    parts(76) = "ytSZSvB+U98V2fwt+Mmp/DPxAk9rK7Q5wQ3TFfbMX7FuhxxlDafSmr+xRoLMQ9p8tfY4zPsLjIWm0fLYPKcRhJ3imd58G/2kPDnj7TYUlv0+24AKZ716F4y8EaX8QtFe3mRHVx97HNeJeF/2S7LwVefadKhZJM7utfQfhHTbzTdPSK6/1g4r80xkqVGr7TDyPvcNSnWp"
    parts(77) = "8lZHwP8AGD9jO40G6ub7QIJLlmJYrjIr531Hwb4j8PzMt/DJZ7T1UEV+y89us6sjqpU9ciuA8TfA3wx4sZmvbMOW64Ar6LB8Szp0+Sqro+exfD6lPmpM/LbQ/ipq/hNwLa5mkK/3ia9H0H9tvxdoKiOKHzMcfMK+o/Ff7F/h2ZmbTbPBavNNa/YlvfmNlbAN2r3Fjsux"
    parts(78) = "kP3jSuePLA42g/dTOEf9vjxfIvz2yKfYVz2tftneLdcjaNotoPcV3D/sR+JZGObcVLbfsS+Io2G+34p0XlNN3UkRKOPktUz528QfE/V/ETs9xczRlj0DGsGO5kvJMPNJJn+8Sa+zdH/YqmLr9ttuO9el+Gf2NfDFvJG13aZx1rsqZzgsLG9JpnMsrxWJladz8/7DwPr2"
    parts(79) = "tTKul2jTbj2WvcPhl+yLrPieaKTXLOS3jOMnFfdfhn4GeF/CpVrO0VCvTIFd5DaLbxhEVQg4GBXzeM4srVY8tNWPo8Hw3Gm+aZ5V8KfgLonw2tYvs0ayyqOrjNeqKoTAUAD0FTGL5fl60yOFw2Wr4itiJ4iTlUep9jh8LGhHlSPEv2pNcXR/ANwS+09K/L2+vUm1i7mY"
    parts(80) = "/fYkcV+uvxY+G0HxA0Z7OVN4Y14fJ+xboLbG+yfN3r7jIc2p4GnyyZ8fnGVzxUtD4b+F9iviHx3a2Kje7EHbiv1r+G9j/Zfg7TLZxtMcQGK8Y8D/ALJeheEvEkWrQWu2dOhxX0PaWYhtUiA4UcVwcQZrHMHFR6HdkuWvB7oe7DdtbgVwnxY+G9n8RvDtxZ3KA7Yz5Zx3"
    parts(81) = "xXeLFuGZOTR5fI4+XtXyVKrKlJTj0PqsRRjWjyyPiv4I/sdQaB4ml1TUYmUxzEorjgjNfZGnabBpdqkMCKiKMDaKvlM4GBt70zyTuP8Ad7V0YvHVsTbmZy4fA0qOqRDM628LSOcYr8xv2xPFS6p8Q76zL5VWzX6d3Nr9pt2R+9fOfjz9lHSPGXiafU7i33tJ1Nepk+Kh"
    parts(82) = "hayqzPLzjDyxEOWJ+a0N5HtVd3QY6V1/wTsTr3xIhtlBZAQelfb6/sU+Hthxac10Pw5/ZP0bwX4iGpRWwVx3r73GcQ0atNpSPisHkdSFVScT27wZafYfDOnRYxtjA/St/hqhhtBb28cSjCqMCpo0PWvySpLnm5dz9Uox5KaiSEbVpEobNKq4rM2HUlLRQAgzRS0UAFFF"
    parts(83) = "FABRRRQAUlLSUAB4pvBOcc0u2jb3o0AY+c5o3A05lz3pu35s1N+w9CNmbdgU/hV+YZNO207Z6809RNRI+JOnFNaNl5zxUjx5HBxQIz6072J5IkHm84xzUqr3PSn+WKNvvQpD5YiDHXHFOK9xxRwKTbuPWhpMYUlO20baSSQmgyKT5snJyKXbkdaVV2j1o1BCc44pBnJz"
    parts(84) = "ThxQRmgZHht2c8U5WycUEUVQtQ3cGj7w55pfel6igYzdgH0pd25cil20BfyrPUEA5Wk2nbS7aXpV6hZDPalPTNFLjimStxFx1xR16Uu3gUMPlpFCA54pNvY80tFApCKewFIWIbBp1IVpj2FZtoyeaTrjFG3bweaWgdrg3t1o3BQO5oYfLRWeqFoJ7GhlP0FOxxRVa2DY"
    parts(85) = "RR8vNFH8VL1pq/UWkgWk6NxRjJpNvzdaEMXPam7hu206lOM9OaUrvYNAztFC880h5xRQkDGuPmFOUDpRSqO9UKyGyMV6VFIx7VYOO9MKg1PNbYJaqxW8uTqDTt23huTUoWk2BTzzQpa6mUaYzbupwUinn5qFFaPUtUorUX+Hml6rQV3ClUEVnsVqC8LSMSKceaNtP1GJ"
    parts(86) = "97mmjvTtuKRV60tVsFkxFIYUv0pelJQr7sl6bCZ7GlpdtIVqrlLYI88g0K3zkdqNvINOx3qeoDD1xQy54FOVdue9IV3UJuwuVBxj3oyaMdqVF2rzTK0QfWlHFLRQIaOpp1FFABRRRQAUUUUAFFFFABRRRQAUUUUAFI1Bz2ox60rANopT7UlOwrhT6aBS0ajDFG0Ud6Wk"
    parts(87) = "AnQU2nGkzU8oCUUUorTZai5WL2pDmlNJuqN9hiUUUvUVVrBcB1oNJSjrQwe4UnNFOGaVxi0lI2cU0ZoWoWHNSrSUtUICaTJpcU1uKV0Av4UlG7NOwKQCbqXqKTbSilzDDbRtpaSncQmOaDRtp1StXqAylx70cUHiqbsGoNSUp9aaKpaodh4o60LSDvSEJSrRtNJQxjsU"
    parts(88) = "hWgGlzU3tuIbS8etHWm96L9h2uLRS9KMU+YloOtB4peBR1FMYgFHFFJULVgFIetO4FHFNoLiKtFLto2miL7g7huNOpuDSim2hK4Z5oJpG60ooGB6UgozSUwFPWl6UmaGpXAF60tNWloY2L1opFpaFsIWk3UU0ilsIXJpc0gpDVXQK4+ikWloGFFFFABRRRQAUUUUAFFF"
    parts(89) = "FABRRRQAUUUUAFFFFACcUUtFJgJS0UUtQCiiimAUmKWimAm0Um2nUUAIxo4opaAGsKKdRQIZQvWnYooeoPcTrQ1OpMUDGDrzTttLiloEroQCkzzTqKBjc0dqXFLQA1aXpRS0CEpaKKBiNSAmnUlAC01qdSUANoanbaMUPUGNop2BQRmnsC0EWloHFLUgN3UUtGKEIbTh"
    parts(90) = "S0UxiUnGM06koASg0tB5oAbSn0pQKKAY2nAYoxS0AI1ItOooAKZk0+kpJWAbk05aTbSgYoauMKMdaWigQyl206imA00DpTqKlIBg60u2nUVQCGm5NPpKAAmkWnUUAIeKBS0UAJS0UUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUU"
    parts(91) = "UUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFA"
    parts(92) = "BRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUU"
    parts(93) = "UUAFFFFABRRRQB//2Q=="
    GetBase64 = Join(parts, "")
End Function

