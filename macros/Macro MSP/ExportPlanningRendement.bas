Attribute VB_Name = "Module1"
Sub ExporterTravailParAssignation()
    Dim xlApp As Object
    Dim xlWb As Object
    Dim xlWs As Object
    Dim tsk As Task
    Dim assn As Assignment
    Dim currentDate As Date
    Dim endDate As Date
    Dim row As Integer
    Dim col As Integer
    Dim travailValue As Variant
    
    ' Créer Excel
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = True
    Set xlWb = xlApp.Workbooks.Add
    Set xlWs = xlWb.Sheets(1)
    
    ' Définir les dates
    currentDate = ActiveProject.ProjectStart
    endDate = ActiveProject.ProjectFinish
    
    ' En-têtes des colonnes (dates)
    col = 2
    Do While currentDate <= endDate
        xlWs.Cells(1, col).Value = Format(currentDate, "dd/mm")
        currentDate = currentDate + 1
        col = col + 1
    Loop
    
    ' Parcourir les tâches et leurs assignations
    row = 2
    For Each tsk In ActiveProject.Tasks
        If Not tsk Is Nothing Then
            ' Parcourir chaque assignation de la tâche
            For Each assn In tsk.Assignments
                ' Nom de l'assignation (Tâche - Ressource)
                xlWs.Cells(row, 1).Value = tsk.Name & " - " & assn.ResourceName
                
                currentDate = ActiveProject.ProjectStart
                col = 2
                
                ' ? EXTRAIRE LA LIGNE "TRAVAIL" DE L'ASSIGNATION
                Do While currentDate <= endDate
                    On Error Resume Next
                    
                    ' TimeScaleData "Work" pour cette assignation spécifique
                    travailValue = assn.TimeScaleData(currentDate, currentDate + 1, pjAssignmentTimescaledWork, pjTimescaleDays).Item(1).Value
                    
                    If Err.Number = 0 And Not IsEmpty(travailValue) And travailValue > 0 Then
                        ' Données brutes de la ligne "Travail" par assignation
                        xlWs.Cells(row, col).Value = travailValue
                    Else
                        xlWs.Cells(row, col).Value = 0
                    End If
                    
                    Err.Clear
                    currentDate = currentDate + 1
                    col = col + 1
                Loop
                
                row = row + 1
            Next assn
        End If
    Next tsk
    
    MsgBox "Export 'Travail' par assignation terminé !"
End Sub
