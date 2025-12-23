Option Explicit

' ============================================
' MACRO MS PROJECT - EXPORT JSON PONTIVA
' ============================================
' Exporte un projet MS Project au format JSON
' compatible avec le Dashboard Pontiva
'
' Version: 0.2
' Date: 2025-12-05
' ============================================

' Point d'entrée principal
Sub ExportProjectToJson()
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
    fileName = "Pontiva_" & projectNameClean & "_" & Format(Now, "yyyymmdd_hhnnss") & ".json"
    
    ' Construire le chemin complet
    filePath = downloadsPath & "\" & fileName
    
    json = BuildProjectJson(ActiveProject)
    
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

' Construit le JSON complet pour un projet
Private Function BuildProjectJson(prj As Project) As String
    Dim sb As String
    Dim t As Task
    Dim firstTask As Boolean
    
    sb = ""
    sb = sb & "{" & vbCrLf
    sb = sb & "  ""version"": ""0.2""," & vbCrLf
    sb = sb & "  ""project_name"": """ & JsonEscape(prj.Name) & """," & vbCrLf
    sb = sb & "  ""export_date"": """ & Format(Date, "yyyy-mm-dd") & """," & vbCrLf
    sb = sb & "  ""tasks"": [" & vbCrLf
    
    firstTask = True
    
    For Each t In prj.Tasks
        If Not t Is Nothing Then
            ' On exporte uniquement les tâches non récapitulatives, avec un nom
            If t.Summary = False And Trim(t.Name) <> "" Then
                If Not firstTask Then
                    sb = sb & "," & vbCrLf
                End If
                sb = sb & BuildTaskJsonNode(t)
                firstTask = False
            End If
        End If
    Next t
    
    sb = sb & vbCrLf & "  ]" & vbCrLf
    sb = sb & "}" & vbCrLf
    
    BuildProjectJson = sb
End Function

' Construit le JSON pour une tâche complète
Private Function BuildTaskJsonNode(t As Task) As String
    Dim sb As String
    Dim status As String
    
    ' Déterminer le statut de la tâche
    status = DetermineTaskStatus(t)
    
    sb = "    {" & vbCrLf
    sb = sb & "      ""uid"": " & t.UniqueID & "," & vbCrLf
    sb = sb & "      ""name"": """ & JsonEscape(t.Name) & """," & vbCrLf
    sb = sb & "      ""duration"": """ & FormatDuration(t.Duration) & """," & vbCrLf
    sb = sb & "      ""start"": """ & FormatDateISO(t.Start) & """," & vbCrLf
    sb = sb & "      ""finish"": """ & FormatDateISO(t.Finish) & """," & vbCrLf
    sb = sb & "      ""predecessors"": """ & JsonEscape(GetPredecessors(t)) & """," & vbCrLf
    sb = sb & "      ""successors"": """ & JsonEscape(GetSuccessors(t)) & """," & vbCrLf
    sb = sb & "      ""resources"": [" & vbCrLf
    sb = sb & BuildTaskResourcesArray(t)
    sb = sb & vbCrLf & "      ]," & vbCrLf
    sb = sb & "      ""percent_complete"": " & FormatNumberDot(t.PercentComplete) & "," & vbCrLf
    sb = sb & "      ""physical_percent_complete"": " & FormatNumberDot(t.PhysicalPercentComplete) & "," & vbCrLf
    sb = sb & "      ""percent_work_complete"": " & FormatNumberDot(t.PercentWorkComplete) & "," & vbCrLf
    sb = sb & "      ""constraint_date"": """ & FormatDateISO(t.ConstraintDate) & """," & vbCrLf
    sb = sb & "      ""baseline_start"": """ & FormatDateISO(t.BaselineStart) & """," & vbCrLf
    sb = sb & "      ""baseline_finish"": """ & FormatDateISO(t.BaselineFinish) & """," & vbCrLf
    sb = sb & "      ""baseline_duration"": """ & FormatDuration(t.BaselineDuration) & """," & vbCrLf
    sb = sb & "      ""planned_duration"": """ & FormatDuration(t.Duration) & """," & vbCrLf
    sb = sb & "      ""actual_duration"": """ & FormatDuration(t.ActualDuration) & """," & vbCrLf
    sb = sb & "      ""remaining_duration"": """ & FormatDuration(t.RemainingDuration) & """," & vbCrLf
    sb = sb & "      ""status"": """ & status & """," & vbCrLf
    sb = sb & "      ""scheduled_finish"": """ & FormatDateISO(t.Finish) & """," & vbCrLf
    sb = sb & "      ""actual_finish"": """ & FormatDateISO(t.ActualFinish) & """," & vbCrLf
    sb = sb & "      ""notes"": """ & JsonEscape(GetTaskNotes(t)) & """," & vbCrLf
    sb = sb & "      ""actual_work_hours"": " & FormatNumberDot(MinutesToHours(t.ActualWork)) & "," & vbCrLf
    sb = sb & "      ""remaining_work_hours"": " & FormatNumberDot(MinutesToHours(t.RemainingWork)) & "," & vbCrLf
    sb = sb & "      ""custom_fields"": {" & vbCrLf
    sb = sb & BuildCustomFields(t)
    sb = sb & vbCrLf & "      }" & vbCrLf
    sb = sb & "    }"
    
    BuildTaskJsonNode = sb
End Function

' Construit le tableau resources pour une tâche avec tous les détails
Private Function BuildTaskResourcesArray(t As Task) As String
    Dim sb As String
    Dim a As Assignment
    Dim firstRes As Boolean
    Dim resType As String
    
    sb = ""
    firstRes = True
    
    For Each a In t.Assignments
        If Not a Is Nothing Then
            If Not a.Resource Is Nothing Then
                ' Déterminer le type de ressource
                resType = GetResourceType(a.Resource)
                
                If Not firstRes Then
                    sb = sb & "," & vbCrLf
                End If
                sb = sb & "        {" & vbCrLf
                sb = sb & "          ""type"": """ & resType & """," & vbCrLf
                sb = sb & "          ""name"": """ & JsonEscape(a.Resource.Name) & """," & vbCrLf
                sb = sb & "          ""group"": """ & JsonEscape(GetResourceGroup(a.Resource)) & """," & vbCrLf
                sb = sb & "          ""actual_work_hours"": " & FormatNumberDot(MinutesToHours(a.ActualWork)) & "," & vbCrLf
                sb = sb & "          ""remaining_work_hours"": " & FormatNumberDot(MinutesToHours(a.RemainingWork)) & vbCrLf
                sb = sb & "        }"
                firstRes = False
            End If
        End If
    Next a
    
    BuildTaskResourcesArray = sb
End Function

' Construit les champs personnalisés d'une tâche
Private Function BuildCustomFields(t As Task) As String
    Dim sb As String
    Dim fieldList As String
    Dim fieldName As String
    Dim fieldValue As String
    Dim firstField As Boolean
    Dim i As Integer
    
    ' Liste des champs personnalisés à vérifier (Text1-30, Number1-20, Flag1-20, Date1-10)
    fieldList = "Text1,Text2,Text3,Text4,Text5,Text6,Text7,Text8,Text9,Text10,Text11,Text12,Text13,Text14,Text15,Text16,Text17,Text18,Text19,Text20,Text21,Text22,Text23,Text24,Text25,Text26,Text27,Text28,Text29,Text30," & _
                "Number1,Number2,Number3,Number4,Number5,Number6,Number7,Number8,Number9,Number10,Number11,Number12,Number13,Number14,Number15,Number16,Number17,Number18,Number19,Number20," & _
                "Flag1,Flag2,Flag3,Flag4,Flag5,Flag6,Flag7,Flag8,Flag9,Flag10,Flag11,Flag12,Flag13,Flag14,Flag15,Flag16,Flag17,Flag18,Flag19,Flag20," & _
                "Date1,Date2,Date3,Date4,Date5,Date6,Date7,Date8,Date9,Date10"
    
    sb = ""
    firstField = True
    
    ' Parcourir les champs personnalisés
    Dim fields() As String
    fields = Split(fieldList, ",")
    
    For i = 0 To UBound(fields)
        fieldName = Trim(fields(i))
        fieldValue = GetCustomFieldValue(t, fieldName)
        
        ' N'inclure que les champs non vides
        If fieldValue <> "" And fieldValue <> "0" And fieldValue <> "False" And fieldValue <> "01/01/1984" Then
            If Not firstField Then
                sb = sb & "," & vbCrLf
            End If
            
            ' Déterminer le format selon le type de champ
            If Left(fieldName, 4) = "Text" Then
                sb = sb & "        """ & fieldName & """: """ & JsonEscape(fieldValue) & """"
            ElseIf Left(fieldName, 6) = "Number" Then
                sb = sb & "        """ & fieldName & """: " & FormatNumberDot(CDbl(Val(fieldValue)))
            ElseIf Left(fieldName, 4) = "Flag" Then
                sb = sb & "        """ & fieldName & """: " & LCase(fieldValue)
            ElseIf Left(fieldName, 4) = "Date" Then
                sb = sb & "        """ & fieldName & """: """ & FormatDateISO(fieldValue) & """"
            End If
            
            firstField = False
        End If
    Next i
    
    BuildCustomFields = sb
End Function

' Récupère la valeur d'un champ personnalisé
Private Function GetCustomFieldValue(t As Task, fieldName As String) As String
    Dim fieldValue As Variant
    Dim fieldObj As Object
    On Error Resume Next
    
    ' Accéder au champ personnalisé via l'objet Task
    Set fieldObj = t.CustomFields(fieldName)
    
    If Err.Number <> 0 Then
        GetCustomFieldValue = ""
        Err.Clear
        Exit Function
    End If
    
    fieldValue = fieldObj.Value
    
    If Err.Number <> 0 Then
        GetCustomFieldValue = ""
        Err.Clear
        Exit Function
    End If
    
    If IsNull(fieldValue) Or IsEmpty(fieldValue) Then
        GetCustomFieldValue = ""
    Else
        GetCustomFieldValue = CStr(fieldValue)
    End If
End Function

' Détermine le statut de la tâche (on_time, late, not_started)
Private Function DetermineTaskStatus(t As Task) As String
    Dim today As Date
    Dim finishDate As Date
    Dim percentComplete As Double
    
    today = Date
    finishDate = t.Finish
    percentComplete = t.PercentComplete
    
    ' Si la tâche est terminée
    If percentComplete = 100 Then
        DetermineTaskStatus = "completed"
    ' Si la tâche n'a pas commencé
    ElseIf percentComplete = 0 And t.Start > today Then
        DetermineTaskStatus = "not_started"
    ' Si la tâche est en retard (fin prévue < aujourd'hui et pas terminée)
    ElseIf finishDate < today And percentComplete < 100 Then
        DetermineTaskStatus = "late"
    ' Sinon à l'heure
    Else
        DetermineTaskStatus = "on_time"
    End If
End Function

' Récupère les prédécesseurs d'une tâche
Private Function GetPredecessors(t As Task) As String
    Dim pred As TaskDependency
    Dim predList As String
    Dim first As Boolean
    On Error Resume Next
    
    predList = ""
    first = True
    
    For Each pred In t.TaskDependencies
        If Err.Number <> 0 Then
            Err.Clear
            Exit For
        End If
        
        ' pred.From est la tâche prédécesseur, pred.To est la tâche successeur
        ' Si pred.To = t, alors pred.From est un prédécesseur de t
        If pred.To.UniqueID = t.UniqueID Then
            If Not first Then
                predList = predList & ", "
            End If
            predList = predList & CStr(pred.From.UniqueID)
            first = False
        End If
    Next pred
    
    GetPredecessors = predList
End Function

' Récupère les successeurs d'une tâche
Private Function GetSuccessors(t As Task) As String
    Dim pred As TaskDependency
    Dim succList As String
    Dim first As Boolean
    On Error Resume Next
    
    succList = ""
    first = True
    
    For Each pred In t.TaskDependencies
        If Err.Number <> 0 Then
            Err.Clear
            Exit For
        End If
        
        ' Si pred.From = t, alors pred.To est un successeur de t
        If pred.From.UniqueID = t.UniqueID Then
            If Not first Then
                succList = succList & ", "
            End If
            succList = succList & CStr(pred.To.UniqueID)
            first = False
        End If
    Next pred
    
    GetSuccessors = succList
End Function

' Récupère le type de ressource (work, material, cost)
Private Function GetResourceType(res As Resource) As String
    On Error Resume Next
    
    If res.Type = pjResourceTypeWork Then
        GetResourceType = "work"
    ElseIf res.Type = pjResourceTypeMaterial Then
        GetResourceType = "material"
    ElseIf res.Type = pjResourceTypeCost Then
        GetResourceType = "cost"
    Else
        GetResourceType = "work" ' Par défaut
    End If
End Function

' Récupère le groupe de ressources
Private Function GetResourceGroup(res As Resource) As String
    On Error Resume Next
    Dim groupValue As String
    
    groupValue = res.Group
    
    If Err.Number <> 0 Or IsNull(groupValue) Or groupValue = "" Then
        GetResourceGroup = ""
        Err.Clear
    Else
        GetResourceGroup = groupValue
    End If
End Function

' Récupère les notes d'une tâche
Private Function GetTaskNotes(t As Task) As String
    On Error Resume Next
    Dim notesValue As String
    
    notesValue = t.Notes
    
    If Err.Number <> 0 Or IsNull(notesValue) Then
        GetTaskNotes = ""
        Err.Clear
    Else
        GetTaskNotes = notesValue
    End If
End Function

' Conversion minutes -> heures
Private Function MinutesToHours(m As Variant) As Double
    If IsNull(m) Or IsEmpty(m) Then
        MinutesToHours = 0#
    Else
        MinutesToHours = CDbl(m) / 60#
    End If
End Function

' Formate une durée en heures (conversion depuis les minutes MSP)
Private Function FormatDuration(duration As Variant) As String
    Dim hours As Double
    On Error Resume Next
    
    If IsNull(duration) Or IsEmpty(duration) Then
        FormatDuration = "0"
        Exit Function
    End If
    
    ' Convertir la durée en heures
    hours = MinutesToHours(duration)
    FormatDuration = FormatNumberDot(hours)
End Function

' Formate une date au format ISO (yyyy-mm-dd)
Private Function FormatDateISO(dateValue As Variant) As String
    On Error Resume Next
    
    If IsNull(dateValue) Or IsEmpty(dateValue) Then
        FormatDateISO = ""
        Exit Function
    End If
    
    ' Vérifier si c'est une date valide (pas la date par défaut MSP)
    If CDate(dateValue) = CDate("01/01/1984") Or CDate(dateValue) = CDate("31/12/2049") Then
        FormatDateISO = ""
    Else
        FormatDateISO = Format(CDate(dateValue), "yyyy-mm-dd")
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

' Nettoie un nom de fichier en enlevant les caractères invalides
Private Function CleanFileName(fileName As String) As String
    Dim s As String
    s = fileName
    ' Remplacer les caractères invalides par des underscores
    s = Replace(s, "\", "_")
    s = Replace(s, "/", "_")
    s = Replace(s, ":", "_")
    s = Replace(s, "*", "_")
    s = Replace(s, "?", "_")
    s = Replace(s, """", "_")
    s = Replace(s, "<", "_")
    s = Replace(s, ">", "_")
    s = Replace(s, "|", "_")
    ' Limiter la longueur
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
