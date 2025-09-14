Sub Import_BJ_WithHierarchy_Omexom_And_Sove()

    Dim xlApp As Object, xlBook As Object, xlSheet As Object
    Dim pjApp As MSProject.Application, pjProj As MSProject.Project
    Dim feuilles As Variant: feuilles = Array("Omexom", "Sove")
    Dim parentFeuille As Task, tacheParent As Task
    Dim i As Long, rowOffset As Long: rowOffset = 4
    Dim lastRow As Long, feuille As Variant

    ' Liste des 8 tâches en plus à ajouter
    Dim autresTaches As Variant
    autresTaches = Array( _
        "MC4 Modules", "MC4 BJ", "Pose BJ", _
        "Remontées sous BJ (75 +80)", "Test Pivostest", _
        "Raccordement de la boite de jonction côté solaire y compris arrivée des câbles et remontées", _
        "Test isolement Continuité DC", "Travaux finitions" _
    )

    ' Ouvre Excel
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = False
    Set xlBook = xlApp.Workbooks.Open("C:\Users\Antoi\Downloads\BJ_Omexom_Sove.xlsx")

    ' Ouvre Project
    Set pjApp = MSProject.Application
    pjApp.Visible = True
    pjApp.FileNew
    Set pjProj = pjApp.ActiveProject

    ' Créer les ressources
    Dim rMonteurs As Resource
    Set rMonteurs = GetOrCreateWorkResource("Monteurs")

    Dim rPose As Resource, rRacc As Resource, rColson As Resource
    Set rPose = GetOrCreateMaterialResource("Pose DC")
    Set rRacc = GetOrCreateMaterialResource("RACC Modules")
    Set rColson = GetOrCreateMaterialResource("Colsonnage Modules")

    ' Parcours des feuilles (Omexom et Sove)
    For Each feuille In feuilles
        Set xlSheet = xlBook.Sheets(feuille)
        lastRow = xlSheet.Cells(xlSheet.Rows.Count, "B").End(-4162).Row

        ' Tâche racine
        Set parentFeuille = pjProj.Tasks.Add("[" & feuille & "]")
        parentFeuille.OutlineLevel = 1

        For i = rowOffset To lastRow
            Dim onduleurID As Variant, bjID As Variant
            Dim qPose As Double, qRacc As Double, qCols As Double
            Dim hPose As Double, hRacc As Double, hCols As Double
            Dim tPose As Task, tRacc As Task, tCols As Task, tAutre As Task
            Dim a As Assignment
            Dim tacheName As String

            onduleurID = xlSheet.Cells(i, "A").Value
            bjID = xlSheet.Cells(i, "B").Value
            If IsEmpty(onduleurID) Or IsEmpty(bjID) Then GoTo NextBJ

            ' Quantités
            qPose = xlSheet.Cells(i, "E").Value
            qRacc = xlSheet.Cells(i, "F").Value
            qCols = xlSheet.Cells(i, "G").Value

            ' Heures (que pour Omexom)
            If feuille = "Omexom" Then
                hPose = xlSheet.Cells(i, "P").Value
                hRacc = xlSheet.Cells(i, "Q").Value
                hCols = xlSheet.Cells(i, "R").Value
            Else
                hPose = 0: hRacc = 0: hCols = 0
            End If

            ' Tâche BJ
            Set tacheParent = pjProj.Tasks.Add("OND" & onduleurID & " – BJ" & bjID)
            tacheParent.OutlineLevel = parentFeuille.OutlineLevel + 1

            ' Pose DC
            Set tPose = pjProj.Tasks.Add("OND" & onduleurID & " – BJ" & bjID & " – Pose DC")
            tPose.OutlineLevel = tacheParent.OutlineLevel + 1
            tPose.Text1 = "Pose DC": tPose.Text2 = feuille
            If hPose > 0 Then
                Set a = tPose.Assignments.Add(ResourceID:=rMonteurs.ID)
                a.Work = hPose * 60
            End If
            If qPose > 0 Then
                Set a = tPose.Assignments.Add(ResourceID:=rPose.ID)
                a.Units = qPose
            End If

            ' RACC Modules
            Set tRacc = pjProj.Tasks.Add("OND" & onduleurID & " – BJ" & bjID & " – RACC Modules")
            tRacc.OutlineLevel = tacheParent.OutlineLevel + 1
            tRacc.Text1 = "RACC Modules": tRacc.Text2 = feuille
            If hRacc > 0 Then
                Set a = tRacc.Assignments.Add(ResourceID:=rMonteurs.ID)
                a.Work = hRacc * 60
            End If
            If qRacc > 0 Then
                Set a = tRacc.Assignments.Add(ResourceID:=rRacc.ID)
                a.Units = qRacc
            End If

            ' Colsonnage Modules
            Set tCols = pjProj.Tasks.Add("OND" & onduleurID & " – BJ" & bjID & " – Colsonnage Modules")
            tCols.OutlineLevel = tacheParent.OutlineLevel + 1
            tCols.Text1 = "Colsonnage Modules": tCols.Text2 = feuille
            If hCols > 0 Then
                Set a = tCols.Assignments.Add(ResourceID:=rMonteurs.ID)
                a.Work = hCols * 60
            End If
            If qCols > 0 Then
                Set a = tCols.Assignments.Add(ResourceID:=rColson.ID)
                a.Units = qCols
            End If

            ' Autres tâches à créer vides
            Dim nomTache As Variant
            For Each nomTache In autresTaches
                Set tAutre = pjProj.Tasks.Add("OND" & onduleurID & " – BJ" & bjID & " – " & nomTache)
                tAutre.OutlineLevel = tacheParent.OutlineLevel + 1
                tAutre.Text1 = nomTache
                tAutre.Text2 = feuille
            Next nomTache

NextBJ:
        Next i
    Next feuille

    MsgBox "Import terminé avec hiérarchie complète.", vbInformation

    xlBook.Close False
    xlApp.Quit
    Set xlSheet = Nothing: Set xlBook = Nothing: Set xlApp = Nothing

End Sub

' Ressource travail
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

' Ressource consommable
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




