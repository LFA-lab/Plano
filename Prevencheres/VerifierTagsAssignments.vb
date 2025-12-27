' ======================================================================
' MACRO DE VERIFICATION - EXPORT EXCEL RAPIDE
' ======================================================================
' Exporte tous les assignments avec leurs tags dans Excel pour vérification
' ======================================================================

Sub ExporterAssignmentsAvecTags()
    
    Dim pjApp As MSProject.Application
    Dim pjProj As MSProject.Project
    Dim t As Task
    Dim a As Assignment
    
    Set pjApp = MSProject.Application
    Set pjProj = pjApp.ActiveProject
    
    If pjProj Is Nothing Then
        MsgBox "Aucun projet MS Project ouvert.", vbCritical
        Exit Sub
    End If
    
    ' Création Excel
    Dim xlApp As Object
    Dim xlBook As Object
    Dim xlSheet As Object
    
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = True
    Set xlBook = xlApp.Workbooks.Add
    Set xlSheet = xlBook.Sheets(1)
    
    ' En-têtes
    xlSheet.Cells(1, 1).Value = "Tâche"
    xlSheet.Cells(1, 2).Value = "Ressource"
    xlSheet.Cells(1, 3).Value = "Type Ressource"
    xlSheet.Cells(1, 4).Value = "Unités"
    xlSheet.Cells(1, 5).Value = "Tranche (Assignment)"
    xlSheet.Cells(1, 6).Value = "Zone (Assignment)"
    xlSheet.Cells(1, 7).Value = "Sous-Zone (Assignment)"
    xlSheet.Cells(1, 8).Value = "Type (Assignment)"
    xlSheet.Cells(1, 9).Value = "Entreprise (Assignment)"
    
    With xlSheet.Range("A1:I1")
        .Font.Bold = True
        .Interior.Color = RGB(200, 220, 255)
    End With
    
    Dim ligneExcel As Long
    ligneExcel = 2
    
    ' Parcours complet
    For Each t In pjProj.Tasks
        If Not t Is Nothing Then
            If Not t.Summary Then
                
                For Each a In t.Assignments
                    If Not a Is Nothing Then
                        
                        xlSheet.Cells(ligneExcel, 1).Value = t.Name
                        xlSheet.Cells(ligneExcel, 2).Value = a.ResourceName
                        
                        ' Type ressource
                        Dim typeRessource As String
                        Select Case a.Resource.Type
                            Case pjResourceTypeWork
                                typeRessource = "Travail"
                            Case pjResourceTypeMaterial
                                typeRessource = "Matériel/Consommable"
                            Case pjResourceTypeCost
                                typeRessource = "Coût"
                            Case Else
                                typeRessource = "Inconnu"
                        End Select
                        xlSheet.Cells(ligneExcel, 3).Value = typeRessource
                        
                        xlSheet.Cells(ligneExcel, 4).Value = a.Units
                        
                        ' Tags hérités (c'est ici qu'on vérifie !)
                        xlSheet.Cells(ligneExcel, 5).Value = a.Text1  ' Tranche
                        xlSheet.Cells(ligneExcel, 6).Value = a.Text2  ' Zone
                        xlSheet.Cells(ligneExcel, 7).Value = a.Text3  ' Sous-Zone
                        xlSheet.Cells(ligneExcel, 8).Value = a.Text4  ' Type
                        xlSheet.Cells(ligneExcel, 9).Value = a.Text5  ' Entreprise
                        
                        ' Coloration selon type
                        If typeRessource = "Matériel/Consommable" Then
                            xlSheet.Range("A" & ligneExcel & ":I" & ligneExcel).Interior.Color = RGB(255, 255, 200)
                        ElseIf typeRessource = "Travail" Then
                            xlSheet.Range("A" & ligneExcel & ":I" & ligneExcel).Interior.Color = RGB(200, 255, 200)
                        End If
                        
                        ligneExcel = ligneExcel + 1
                        
                    End If
                Next a
                
            End If
        End If
    Next t
    
    ' Ajustement colonnes
    xlSheet.Columns("A:I").AutoFit
    
    ' Ajouter un filtre automatique
    xlSheet.Range("A1:I1").AutoFilter
    
    MsgBox "Export terminé !" & vbCrLf & vbCrLf & _
           (ligneExcel - 2) & " affectation(s) exportée(s)." & vbCrLf & vbCrLf & _
           "Vérifie les colonnes E à I pour les tags des assignments." & vbCrLf & vbCrLf & _
           "Légende couleurs:" & vbCrLf & _
           "  - Jaune = Ressources Matériel/Consommables" & vbCrLf & _
           "  - Vert = Ressources Travail (Monteurs)", vbInformation
    
    Set xlSheet = Nothing
    Set xlBook = Nothing
    Set xlApp = Nothing
    
End Sub

