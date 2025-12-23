Option Explicit

' ============================================
' MACRO MS PROJECT - EXPORT JSON SUIVI MECA/ELEC
' ============================================
' Exporte le suivi mécanique et électrique au format JSON
' Compatible avec le Dashboard Pontiva
'
' Version: 0.1
' Date: 2025-12-05
' ============================================

' Point d'entrée principal
Public Sub ExportSuiviMecaElecJson()
    Dim json As String
    Dim filePath As String
    Dim downloadsPath As String
    Dim fileName As String
    Dim projectNameClean As String
    
    If ActiveProject Is Nothing Then
        MsgBox "Aucun projet actif.", vbExclamation
        Exit Sub
    End If
    
    ' Récupérer le chemin du dossier Téléchargements
    downloadsPath = Environ$("USERPROFILE") & "\Downloads"
    
    ' Vérifier que le dossier existe, sinon utiliser le Bureau
    If Dir(downloadsPath, vbDirectory) = "" Then
        downloadsPath = Environ$("USERPROFILE") & "\Desktop"
    End If
    
    ' Nettoyer le nom du projet pour créer un nom de fichier valide
    projectNameClean = CleanFileName(ActiveProject.Name)
    
    ' Générer un nom de fichier avec date et heure
    fileName = "Suivi_MecaElec_" & projectNameClean & "_" & Format(Now, "yyyymmdd_hhnnss") & ".json"
    
    ' Construire le chemin complet
    filePath = downloadsPath & "\" & fileName
    
    json = BuildSuiviJson(ActiveProject)
    
    If Len(json) = 0 Then
        MsgBox "Rien à exporter.", vbExclamation
        Exit Sub
    End If
    
    If WriteTextFile(filePath, json) Then
        MsgBox "Export JSON terminé !" & vbCrLf & vbCrLf & "Fichier enregistré dans :" & vbCrLf & filePath, vbInformation
    Else
        MsgBox "Erreur lors de l'écriture du fichier JSON.", vbCritical
    End If
End Sub

' Construit le JSON complet pour le suivi Meca/Elec
Private Function BuildSuiviJson(prj As Project) As String
    Dim sb As String
    Dim startDate As Date, endDate As Date
    
    startDate = prj.ProjectStart
    endDate = prj.ProjectFinish
    
    sb = ""
    sb = sb & "{" & vbCrLf
    sb = sb & "  ""version"": ""suivi-meca-elec-0.1""," & vbCrLf
    sb = sb & "  ""project_name"": """ & JsonEscape(prj.Name) & """," & vbCrLf
    sb = sb & "  ""export_date"": """ & Format(Date, "yyyy-mm-dd") & """," & vbCrLf
    sb = sb & "  ""mechanical"": {" & vbCrLf
    sb = sb & BuildGroupData(prj, "Mecanique", startDate, endDate)
    sb = sb & "  }," & vbCrLf
    sb = sb & "  ""electrical"": {" & vbCrLf
    sb = sb & BuildGroupData(prj, "Electrique", startDate, endDate)
    sb = sb & "  }" & vbCrLf
    sb = sb & "}" & vbCrLf
    
    BuildSuiviJson = sb
End Function

' Construit les données pour un groupe donné (Mecanique ou Electrique)
Private Function BuildGroupData(prj As Project, groupName As String, startDate As Date, endDate As Date) As String
    Dim sb As String
    Dim resList As Collection
    Dim resAssignments As Object
    Dim totalPlanned As Object
    Dim dailyActual As Object
    Dim cumActual As Object
    Dim datesAsc As Variant
    
    ' Collecter les ressources du groupe
    Set resList = GetSortedResourcesByGroup(prj, groupName)
    
    If resList.Count = 0 Then
        ' Groupe vide
        BuildGroupData = "    ""resources"": []," & vbCrLf & _
                        "    ""dates"": []," & vbCrLf & _
                        "    ""daily"": []," & vbCrLf & _
                        "    ""global"": {" & vbCrLf & _
                        "      ""total_planned_work_hours"": 0," & vbCrLf & _
                        "      ""total_actual_work_hours"": 0," & vbCrLf & _
                        "      ""progress_percent"": 0" & vbCrLf & _
                        "    }"
        Exit Function
    End If
    
    ' Mapper les assignations
    Set resAssignments = MapAssignmentsByResource(resList)
    
    ' Calculer les données
    Set totalPlanned = ComputeTotalPlannedWork(resAssignments)
    Set dailyActual = ComputeDailyActualWork(resAssignments, startDate, endDate)
    datesAsc = BuildActualDatesIndex(dailyActual, True)
    Set cumActual = ComputeCumulativeActual(dailyActual, datesAsc)
    
    ' Construire le JSON
    sb = ""
    
    ' Liste des ressources avec récap
    sb = sb & "    ""resources"": [" & vbCrLf
    sb = sb & BuildResourcesList(resList, totalPlanned, cumActual, datesAsc)
    sb = sb & "    ]," & vbCrLf
    
    ' Liste des dates
    sb = sb & "    ""dates"": [" & vbCrLf
    sb = sb & BuildDatesList(datesAsc)
    sb = sb & "    ]," & vbCrLf
    
    ' Données quotidiennes
    sb = sb & "    ""daily"": [" & vbCrLf
    sb = sb & BuildDailyData(datesAsc, resList, totalPlanned, dailyActual, cumActual)
    sb = sb & "    ]," & vbCrLf
    
    ' Récap global
    sb = sb & "    ""global"": {" & vbCrLf
    sb = sb & BuildGlobalRecap(resList, totalPlanned, cumActual, datesAsc)
    sb = sb & "    }"
    
    BuildGroupData = sb
End Function

' Collecte les ressources d'un groupe donné, triées par ID de tâche
Private Function GetSortedResourcesByGroup(prj As Project, groupName As String) As Collection
    On Error GoTo ErrorHandler
    
    Dim resArray() As Variant
    Dim resCount As Long
    resCount = 0
    ReDim resArray(0 To 50)
    
    Dim res As Resource
    For Each res In prj.Resources
        If Not res Is Nothing Then
            ' Vérifier si la ressource appartient au groupe
            If UCase(Trim(res.Group)) <> UCase(groupName) Then
                GoTo NextResource
            End If
            
            ' Trouver l'ID minimum de tâche
            Dim minTaskId As Long
            minTaskId = 2147483647
            
            If res.Assignments.Count > 0 Then
                Dim assn As Assignment
                For Each assn In res.Assignments
                    If assn.Task.ID < minTaskId Then
                        minTaskId = assn.Task.ID
                    End If
                Next assn
            End If
            
            ' Redimensionner si nécessaire
            If resCount > UBound(resArray) Then
                ReDim Preserve resArray(0 To UBound(resArray) + 50)
            End If
            
            resArray(resCount) = Array(res.Name, minTaskId)
            resCount = resCount + 1
            
NextResource:
        End If
    Next res
    
    ' Redimensionner à la taille exacte
    If resCount > 0 Then
        ReDim Preserve resArray(0 To resCount - 1)
        
        ' Trier si nécessaire
        If resCount > 1 Then
            QuickSortResources resArray, 0, resCount - 1
        End If
    End If
    
    ' Créer la collection
    Dim sortedResList As Collection
    Set sortedResList = New Collection
    
    Dim i As Long
    For i = 0 To resCount - 1
        sortedResList.Add resArray(i)(0)
    Next i
    
    Set GetSortedResourcesByGroup = sortedResList
    Exit Function
    
ErrorHandler:
    Set sortedResList = New Collection
    Set GetSortedResourcesByGroup = sortedResList
End Function

' Quicksort pour l'array de ressources
Private Sub QuickSortResources(arr As Variant, first As Long, last As Long)
    Dim low As Long, high As Long
    Dim pivot As Long
    Dim temp As Variant
    
    low = first
    high = last
    pivot = arr((first + last) \ 2)(1)
    
    Do While low <= high
        Do While arr(low)(1) < pivot
            low = low + 1
        Loop
        Do While arr(high)(1) > pivot
            high = high - 1
        Loop
        
        If low <= high Then
            temp = arr(low)
            arr(low) = arr(high)
            arr(high) = temp
            low = low + 1
            high = high - 1
        End If
    Loop
    
    If first < high Then QuickSortResources arr, first, high
    If low < last Then QuickSortResources arr, low, last
End Sub

' Index assignations par ressource
Private Function MapAssignmentsByResource(resList As Collection) As Object
    Dim resAssignments As Object
    On Error GoTo DictError
    Set resAssignments = CreateObject("Scripting.Dictionary")
    On Error GoTo 0
    
    Dim resName As Variant
    For Each resName In resList
        Set resAssignments(resName) = New Collection
    Next
    
    Dim res As Resource, assn As Assignment
    For Each res In ActiveProject.Resources
        If Not res Is Nothing And resAssignments.exists(res.Name) Then
            For Each assn In res.Assignments
                resAssignments(res.Name).Add assn
            Next
        End If
    Next
    
    Set MapAssignmentsByResource = resAssignments
    Exit Function
    
DictError:
    Set resAssignments = CreateObject("Scripting.Dictionary")
    Set MapAssignmentsByResource = resAssignments
End Function

' Totaux prévus par ressource (Work en heures)
Private Function ComputeTotalPlannedWork(resAssignments As Object) As Object
    Dim totalPlanned As Object
    Set totalPlanned = CreateObject("Scripting.Dictionary")
    
    Dim resName As Variant, assn As Assignment
    For Each resName In resAssignments.Keys
        Dim totalWork As Double: totalWork = 0
        For Each assn In resAssignments(resName)
            totalWork = totalWork + MinutesToHours(assn.Work)
        Next
        totalPlanned(resName) = totalWork
    Next
    
    Set ComputeTotalPlannedWork = totalPlanned
End Function

' Travail réel par jour (en heures)
Private Function ComputeDailyActualWork(resAssignments As Object, startDate As Date, endDate As Date) As Object
    Dim dailyActual As Object
    Set dailyActual = CreateObject("Scripting.Dictionary")
    
    Dim resName As Variant, assn As Assignment
    For Each resName In resAssignments.Keys
        Set dailyActual(resName) = CreateObject("Scripting.Dictionary")
        
        For Each assn In resAssignments(resName)
            Dim tsv As TimeScaleValues
            Set tsv = assn.TimeScaleData(startDate, endDate + 1, pjAssignmentTimescaledActualWork, pjTimescaleDays)
            
            Dim i As Integer
            For i = 1 To tsv.Count
                If Not tsv(i) Is Nothing And IsNumeric(tsv(i).Value) Then
                    If tsv(i).Value <> 0 Then
                        Dim dateKey As String
                        dateKey = Format(tsv(i).startDate, "yyyy-mm-dd")
                        Dim hoursValue As Double
                        hoursValue = MinutesToHours(tsv(i).Value)
                        
                        If dailyActual(resName).exists(dateKey) Then
                            dailyActual(resName)(dateKey) = dailyActual(resName)(dateKey) + hoursValue
                        Else
                            dailyActual(resName)(dateKey) = hoursValue
                        End If
                    End If
                End If
            Next i
        Next
    Next
    
    Set ComputeDailyActualWork = dailyActual
End Function

' Dates où il y a du réel, triées
Private Function BuildActualDatesIndex(dailyActual As Object, Optional ascending As Boolean = True) As Variant
    Dim dateDict As Object
    Set dateDict = CreateObject("Scripting.Dictionary")
    
    Dim resName As Variant, dateKey As Variant
    For Each resName In dailyActual.Keys
        For Each dateKey In dailyActual(resName).Keys
            If Not dateDict.exists(dateKey) Then
                dateDict.Add dateKey, True
            End If
        Next
    Next
    
    If dateDict.Count = 0 Then
        BuildActualDatesIndex = Array()
        Exit Function
    End If
    
    Dim sortedDates As Variant
    sortedDates = dateDict.Keys
    Call QuickSortDates(sortedDates, LBound(sortedDates), UBound(sortedDates))
    
    If Not ascending Then
        sortedDates = ReverseArray(sortedDates)
    End If
    
    BuildActualDatesIndex = sortedDates
End Function

' Tri rapide pour les dates
Private Sub QuickSortDates(arr As Variant, first As Long, last As Long)
    Dim low As Long, high As Long
    Dim mid As String, temp As String
    
    low = first
    high = last
    mid = arr((first + last) \ 2)
    
    Do While low <= high
        Do While arr(low) < mid
            low = low + 1
        Loop
        Do While arr(high) > mid
            high = high - 1
        Loop
        If low <= high Then
            temp = arr(low)
            arr(low) = arr(high)
            arr(high) = temp
            low = low + 1
            high = high - 1
        End If
    Loop
    
    If first < high Then QuickSortDates arr, first, high
    If low < last Then QuickSortDates arr, low, last
End Sub

' Cumul par ressource et par date
Private Function ComputeCumulativeActual(dailyActual As Object, orderedDatesAsc As Variant) As Object
    Dim cumActual As Object
    Set cumActual = CreateObject("Scripting.Dictionary")
    
    Dim resName As Variant
    For Each resName In dailyActual.Keys
        Set cumActual(resName) = CreateObject("Scripting.Dictionary")
        
        Dim cumSum As Double: cumSum = 0
        Dim d As Variant
        For Each d In orderedDatesAsc
            If dailyActual(resName).exists(d) Then
                cumSum = cumSum + dailyActual(resName)(d)
            End If
            cumActual(resName)(d) = cumSum
        Next
    Next
    
    Set ComputeCumulativeActual = cumActual
End Function

' Construit la liste des ressources avec récap
Private Function BuildResourcesList(resList As Collection, totalPlanned As Object, cumActual As Object, datesAsc As Variant) As String
    Dim sb As String
    Dim firstRes As Boolean
    firstRes = True
    
    Dim resName As Variant
    For Each resName In resList
        If Not firstRes Then
            sb = sb & "," & vbCrLf
        End If
        
        Dim plannedHours As Double: plannedHours = 0
        Dim actualHours As Double: actualHours = 0
        
        If totalPlanned.exists(resName) Then
            plannedHours = totalPlanned(resName)
        End If
        
        ' Calculer le total réel (maximum du cumul)
        If Not IsEmpty(datesAsc) And UBound(datesAsc) >= LBound(datesAsc) Then
            Dim lastDate As String: lastDate = datesAsc(UBound(datesAsc))
            If cumActual.exists(resName) And cumActual(resName).exists(lastDate) Then
                actualHours = cumActual(resName)(lastDate)
            End If
        End If
        
        Dim progressPercent As Double: progressPercent = 0
        If plannedHours > 0 Then
            progressPercent = Round((actualHours / plannedHours) * 100, 1)
        End If
        
        sb = sb & "      {" & vbCrLf
        sb = sb & "        ""name"": """ & JsonEscape(resName) & """," & vbCrLf
        sb = sb & "        ""total_planned_work_hours"": " & FormatNumberDot(plannedHours) & "," & vbCrLf
        sb = sb & "        ""total_actual_work_hours"": " & FormatNumberDot(actualHours) & "," & vbCrLf
        sb = sb & "        ""progress_percent"": " & FormatNumberDot(progressPercent) & vbCrLf
        sb = sb & "      }"
        
        firstRes = False
    Next
    
    BuildResourcesList = sb
End Function

' Construit la liste des dates
Private Function BuildDatesList(datesAsc As Variant) As String
    Dim sb As String
    Dim firstDate As Boolean
    firstDate = True
    
    If IsEmpty(datesAsc) Or UBound(datesAsc) < LBound(datesAsc) Then
        BuildDatesList = ""
        Exit Function
    End If
    
    Dim d As Variant
    For Each d In datesAsc
        If Not firstDate Then
            sb = sb & "," & vbCrLf
        End If
        sb = sb & "      """ & d & """"
        firstDate = False
    Next
    
    BuildDatesList = sb
End Function

' Construit les données quotidiennes
Private Function BuildDailyData(datesAsc As Variant, resList As Collection, totalPlanned As Object, dailyActual As Object, cumActual As Object) As String
    Dim sb As String
    Dim firstDate As Boolean
    firstDate = True
    
    If IsEmpty(datesAsc) Or UBound(datesAsc) < LBound(datesAsc) Then
        BuildDailyData = ""
        Exit Function
    End If
    
    Dim d As Variant
    For Each d In datesAsc
        If Not firstDate Then
            sb = sb & "," & vbCrLf
        End If
        
        sb = sb & "      {" & vbCrLf
        sb = sb & "        ""date"": """ & d & """," & vbCrLf
        sb = sb & "        ""resources"": [" & vbCrLf
        sb = sb & BuildDailyResourcesData(d, resList, totalPlanned, dailyActual, cumActual)
        sb = sb & vbCrLf & "        ]" & vbCrLf
        sb = sb & "      }"
        
        firstDate = False
    Next
    
    BuildDailyData = sb
End Function

' Construit les données quotidiennes pour toutes les ressources à une date donnée
Private Function BuildDailyResourcesData(dateKey As String, resList As Collection, totalPlanned As Object, dailyActual As Object, cumActual As Object) As String
    Dim sb As String
    Dim firstRes As Boolean
    firstRes = True
    
    Dim resName As Variant
    For Each resName In resList
        If Not firstRes Then
            sb = sb & "," & vbCrLf
        End If
        
        Dim plannedTotal As Double: plannedTotal = 0
        Dim actualCumulative As Double: actualCumulative = 0
        Dim actualDay As Double: actualDay = 0
        
        If totalPlanned.exists(resName) Then
            plannedTotal = totalPlanned(resName)
        End If
        
        If cumActual.exists(resName) And cumActual(resName).exists(dateKey) Then
            actualCumulative = cumActual(resName)(dateKey)
        End If
        
        If dailyActual.exists(resName) And dailyActual(resName).exists(dateKey) Then
            actualDay = dailyActual(resName)(dateKey)
        End If
        
        Dim progressPercent As Double: progressPercent = 0
        If plannedTotal > 0 Then
            progressPercent = Round((actualCumulative / plannedTotal) * 100, 1)
        End If
        
        sb = sb & "          {" & vbCrLf
        sb = sb & "            ""name"": """ & JsonEscape(resName) & """," & vbCrLf
        sb = sb & "            ""planned_total_hours"": " & FormatNumberDot(plannedTotal) & "," & vbCrLf
        sb = sb & "            ""actual_cumulative_hours"": " & FormatNumberDot(actualCumulative) & "," & vbCrLf
        sb = sb & "            ""actual_day_hours"": " & FormatNumberDot(actualDay) & "," & vbCrLf
        sb = sb & "            ""progress_percent"": " & FormatNumberDot(progressPercent) & vbCrLf
        sb = sb & "          }"
        
        firstRes = False
    Next
    
    BuildDailyResourcesData = sb
End Function

' Construit le récap global
Private Function BuildGlobalRecap(resList As Collection, totalPlanned As Object, cumActual As Object, datesAsc As Variant) As String
    Dim sb As String
    
    Dim totalPlannedHours As Double: totalPlannedHours = 0
    Dim totalActualHours As Double: totalActualHours = 0
    
    Dim resName As Variant
    For Each resName In resList
        If totalPlanned.exists(resName) Then
            totalPlannedHours = totalPlannedHours + totalPlanned(resName)
        End If
        
        ' Calculer le total réel (maximum du cumul)
        If Not IsEmpty(datesAsc) And UBound(datesAsc) >= LBound(datesAsc) Then
            Dim lastDate As String: lastDate = datesAsc(UBound(datesAsc))
            If cumActual.exists(resName) And cumActual(resName).exists(lastDate) Then
                totalActualHours = totalActualHours + cumActual(resName)(lastDate)
            End If
        End If
    Next
    
    Dim globalProgress As Double: globalProgress = 0
    If totalPlannedHours > 0 Then
        globalProgress = Round((totalActualHours / totalPlannedHours) * 100, 1)
    End If
    
    sb = sb & "      ""total_planned_work_hours"": " & FormatNumberDot(totalPlannedHours) & "," & vbCrLf
    sb = sb & "      ""total_actual_work_hours"": " & FormatNumberDot(totalActualHours) & "," & vbCrLf
    sb = sb & "      ""progress_percent"": " & FormatNumberDot(globalProgress)
    
    BuildGlobalRecap = sb
End Function

' Helper pour inverser un array
Private Function ReverseArray(arr As Variant) As Variant
    If IsEmpty(arr) Or UBound(arr) < LBound(arr) Then
        ReverseArray = arr
        Exit Function
    End If
    
    Dim result As Variant
    ReDim result(LBound(arr) To UBound(arr))
    
    Dim i As Long, j As Long
    j = UBound(arr)
    For i = LBound(arr) To UBound(arr)
        result(i) = arr(j)
        j = j - 1
    Next
    
    ReverseArray = result
End Function

' Conversion minutes -> heures
Private Function MinutesToHours(m As Variant) As Double
    If IsNull(m) Or IsEmpty(m) Then
        MinutesToHours = 0#
    Else
        MinutesToHours = CDbl(m) / 60#
    End If
End Function

' Échappe les caractères spéciaux pour JSON
Private Function JsonEscape(text As String) As String
    Dim s As String
    s = text
    s = Replace(s, "\", "\\")
    s = Replace(s, """", "\""")
    s = Replace(s, vbCrLf, "\n")
    s = Replace(s, vbCr, "\n")
    s = Replace(s, vbLf, "\n")
    JsonEscape = s
End Function

' Formate un nombre avec point décimal (JSON)
Private Function FormatNumberDot(value As Double) As String
    Dim s As String
    s = CStr(value)
    s = Replace(s, ",", ".")
    FormatNumberDot = s
End Function

' Nettoie un nom de fichier
Private Function CleanFileName(fileName As String) As String
    Dim s As String
    s = fileName
    s = Replace(s, "\", "_")
    s = Replace(s, "/", "_")
    s = Replace(s, ":", "_")
    s = Replace(s, "*", "_")
    s = Replace(s, "?", "_")
    s = Replace(s, """", "_")
    s = Replace(s, "<", "_")
    s = Replace(s, ">", "_")
    s = Replace(s, "|", "_")
    If Len(s) > 50 Then
        s = Left(s, 50)
    End If
    CleanFileName = s
End Function

' Écrit une chaîne dans un fichier texte
Private Function WriteTextFile(filePath As String, content As String) As Boolean
    Dim fileNum As Integer
    On Error GoTo ErrHandler
    
    fileNum = FreeFile
    Open filePath For Output As #fileNum
    Print #fileNum, content;
    Close #fileNum
    
    WriteTextFile = True
    Exit Function
    
ErrHandler:
    On Error Resume Next
    If fileNum <> 0 Then Close #fileNum
    WriteTextFile = False
End Function

