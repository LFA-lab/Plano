Option Explicit

' =============================================================================
' MACRO D'EXPORT GANTT - Fichier Maître Prévenchères
' =============================================================================

Public Sub ExportGantt_3Semaines()
    ' Export une vue Gantt propre sur 3 semaines pour le rapport hebdomadaire
    ' À lancer depuis le fichier maître avant de générer le rapport
    
    Dim outputPath As String
    Dim startDate As Date
    Dim endDate As Date
    
    On Error GoTo EH
    
    ' Définir le chemin de sortie (même dossier que le planning)
    outputPath = ActiveProject.Path & "\Gantt_Export_3Semaines.png"
    
    Debug.Print "=== Export Gantt 3 semaines ==="
    Debug.Print "Fichier: " & outputPath
    
    ' Configuration de la vue
    ViewApply "Gantt Chart"
    
    ' Période: aujourd'hui + 3 semaines
    startDate = Date
    endDate = Date + 21
    Debug.Print "Période: " & Format(startDate, "dd/mm/yyyy") & " au " & Format(endDate, "dd/mm/yyyy")
    
    ' Échelle de temps: par semaines
    ZoomTimescale 3  ' 3 = semaines
    
    ' Positionner sur aujourd'hui
    GoToDate startDate
    
    ' Filtrage optionnel (décommenter si besoin)
    ' FilterApply "Tâches actives"
    
    ' Ajuster la hauteur des lignes pour meilleure lisibilité
    Application.RowHeight = 16
    
    ' Copier l'image dans le presse-papier
    EditCopyPicture
    Debug.Print "Image copiée dans le presse-papier"
    
    ' Sauvegarder l'image
    If SaveImageFromClipboard(outputPath) Then
        MsgBox "✓ Export Gantt réussi !" & vbCrLf & vbCrLf & outputPath, vbInformation, "Export Gantt"
        Debug.Print "✓ Export réussi"
    Else
        MsgBox "Erreur lors de l'export." & vbCrLf & "Utilisez Ctrl+V dans Paint et sauvegardez manuellement.", vbExclamation
        Debug.Print "✗ Échec - utiliser méthode manuelle"
    End If
    
    Exit Sub
    
EH:
    MsgBox "Erreur: " & Err.Description, vbCritical
    Debug.Print "ERREUR: " & Err.Number & " - " & Err.Description
End Sub

Private Function SaveImageFromClipboard(ByVal filePath As String) As Boolean
    ' Sauvegarde l'image du presse-papier en utilisant un document Word temporaire
    
    Dim wdApp As Object
    Dim doc As Object
    
    On Error GoTo EH
    
    ' Créer Word temporaire
    Set wdApp = CreateObject("Word.Application")
    wdApp.Visible = False
    Set doc = wdApp.Documents.Add
    
    ' Coller l'image
    doc.Range.Paste
    
    If doc.InlineShapes.Count = 0 Then
        doc.Close False
        wdApp.Quit
        SaveImageFromClipboard = False
        Exit Function
    End If
    
    ' Exporter en PNG
    doc.InlineShapes(1).Range.CopyAsPicture
    doc.InlineShapes(1).ConvertToShape
    doc.Shapes(1).Export filePath, 13  ' 13 = PNG
    
    ' Fermer Word
    doc.Close False
    wdApp.Quit
    
    SaveImageFromClipboard = True
    Exit Function
    
EH:
    Debug.Print "Erreur SaveImageFromClipboard: " & Err.Description
    On Error Resume Next
    If Not doc Is Nothing Then doc.Close False
    If Not wdApp Is Nothing Then wdApp.Quit
    On Error GoTo 0
    SaveImageFromClipboard = False
End Function