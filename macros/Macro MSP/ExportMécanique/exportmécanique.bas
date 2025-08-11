' ===================================================================
' MACRO UNIFIÉE : EXPORT MÉCANIQUE - 2 ONGLETS
' ===================================================================
' Onglet 1 : Récapitulatif global des ressources mécaniques
' Onglet 2 : Données détaillées (uniquement dates avec valeurs réelles)
' ===================================================================

' === DÉCLARATIONS POUR DOSSIER TÉLÉCHARGEMENTS ===
Private Declare PtrSafe Function SHGetKnownFolderPath Lib "shell32" _
    (ByRef rfid As GUID, ByVal dwFlags As Long, ByVal hToken As LongPtr, ByRef pszPath As LongPtr) As Long

Private Declare PtrSafe Sub CoTaskMemFree Lib "ole32" (ByVal pv As LongPtr)
Private Declare PtrSafe Function PtrToString Lib "kernel32" Alias "lstrcpyW" (ByVal lpString1 As LongPtr, ByVal lpString2 As LongPtr) As Long

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

' === FONCTION UTILITAIRE : DOSSIER TÉLÉCHARGEMENTS ===
Function GetDownloadsFolder() As String
    Const KF_FLAG_DEFAULT As Long = 0
    Dim pathPtr As LongPtr
    Dim folderPath As String
    Dim folderID As GUID

    ' FOLDERID_Downloads = {374DE290-123F-4565-9164-39C4925E467B}
    With folderID
        .Data1 = &H374DE290
        .Data2 = &H123F
        .Data3 = &H4565
        .Data4(0) = &H91: .Data4(1) = &H64: .Data4(2) = &H39
        .Data4(3) = &HC4: .Data4(4) = &H92: .Data4(5) = &H5E
        .Data4(6) = &H46: .Data4(7) = &H7B
    End With

    If SHGetKnownFolderPath(folderID, KF_FLAG_DEFAULT, 0, pathPtr) = 0 Then
        folderPath = Space$(260)
        Call PtrToString(StrPtr(folderPath), pathPtr)
        folderPath = Left(folderPath, InStr(folderPath, vbNullChar) - 1)
        CoTaskMemFree pathPtr
        GetDownloadsFolder = folderPath
    Else
        GetDownloadsFolder = ""
    End If
End Function

' === MACRO PRINCIPALE ===
Sub ExportMecaniqueComplet()
    Debug.Print "=== DÉBUT EXPORT MÉCANIQUE COMPLET ==="
    
    Dim fileName As String, downloadPath As String
    Dim startDate As Date, endDate As Date
    Dim xlApp As Object, xlBook As Object, xlRecapSheet As Object, xlDetailSheet As Object
    
    ' Collections pour les données
    Dim resList As Collection
    Dim resAssignments As Object
    Dim totalPlanned As Object
    Dim dailyActual As Object
    Dim cumActual As Object
    Dim datesAsc As Variant, datesDesc As Variant
    Dim recapData As Collection

    Set recapData = New Collection

    On Error GoTo ErrorHandler

    ' Récupérer dossier Téléchargements
    downloadPath = GetDownloadsFolder()
    If downloadPath = "" Then
        MsgBox "? Impossible de localiser le dossier Téléchargements.", vbCritical
        Exit Sub
    End If

    fileName = downloadPath & "\Export_Mecanique_Complet_" & Format(Now, "yyyymmdd_hhnnss") & ".xlsx"
    startDate = ActiveProject.ProjectStart
    endDate = ActiveProject.ProjectFinish

    Debug.Print "=== ÉTAPE 1: Collecte et tri des ressources mécaniques ==="
    Set resList = GetSortedMechanicalResources(ActiveProject)
    
    If resList.Count = 0 Then
        MsgBox "?? Aucune ressource du groupe Mécanique trouvée dans le projet.", vbExclamation
        Exit Sub
    End If

    Debug.Print "=== ÉTAPE 2: Collecte des assignations ==="
    Set resAssignments = MapAssignmentsByResource(resList)

    Debug.Print "=== ÉTAPE 3: Calcul des données ==="
    Set totalPlanned = ComputeTotalPlannedWork(resAssignments)
    Set dailyActual = ComputeDailyActualWork(resAssignments, startDate, endDate)
    datesAsc = BuildActualDatesIndex(dailyActual, True)
    Set cumActual = ComputeCumulativeActual(dailyActual, datesAsc)
    datesDesc = ReverseArray(datesAsc)

    Debug.Print "=== ÉTAPE 4: Calcul du récapitulatif global ==="
    
    ' Calculer le récapitulatif global (utilise totalPlanned et cumActual)
    Dim resName As Variant
    For Each resName In resList
        Dim recapTotalWork As Double: recapTotalWork = totalPlanned(resName)
        Dim recapTotalActual As Double: recapTotalActual = 0
        
        ' Calculer le total réel (maximum du cumul)
        If Not IsEmpty(datesAsc) And UBound(datesAsc) >= LBound(datesAsc) Then
            Dim lastDate As String: lastDate = datesAsc(UBound(datesAsc))
            If cumActual(resName).exists(lastDate) Then
                recapTotalActual = cumActual(resName)(lastDate)
            End If
        End If

        Dim recapPercent As Double
        If recapTotalWork > 0 Then
            recapPercent = Round((recapTotalActual / recapTotalWork) * 100, 1)
        End If

        Dim recapLineCalc As Collection
        Set recapLineCalc = New Collection
        recapLineCalc.Add resName
        recapLineCalc.Add Round(recapTotalWork, 0)
        recapLineCalc.Add Round(recapTotalActual, 0)
        recapLineCalc.Add recapPercent
        recapData.Add recapLineCalc
        
        Debug.Print "Récap " & resName & ": " & recapTotalWork & "/" & recapTotalActual & " (" & recapPercent & "%)"
    Next

    Debug.Print "=== ÉTAPE 5: Création du fichier Excel ==="

    ' Créer Excel avec 2 onglets
    Set xlApp = CreateObject("Excel.Application")
    Set xlBook = xlApp.Workbooks.Add
    
    ' Supprimer feuilles par défaut sauf une
    xlApp.DisplayAlerts = False
    Do While xlBook.Worksheets.Count > 1
        xlBook.Worksheets(xlBook.Worksheets.Count).Delete
    Loop
    xlApp.DisplayAlerts = True
    
    ' Créer les 2 onglets
    Set xlRecapSheet = xlBook.Worksheets(1)
    xlRecapSheet.Name = "Récapitulatif"
    Set xlDetailSheet = xlBook.Worksheets.Add
    xlDetailSheet.Name = "Données détaillées"
    xlDetailSheet.Move After:=xlRecapSheet

    Debug.Print "=== ONGLET 1: Écriture du récapitulatif ==="
    
    ' === ONGLET 1 : RÉCAPITULATIF ===
    With xlRecapSheet
        .Cells(1, 1).value = "Ressource"
        .Cells(1, 2).value = "Prévu"
        .Cells(1, 3).value = "Réalisé"
        .Cells(1, 4).value = "Pourcentage"
        
        Dim row As Integer: row = 2
        Dim recapLineDetail As Collection
        Dim globalWork As Double: globalWork = 0
        Dim globalActual As Double: globalActual = 0
        
        For Each recapLineDetail In recapData
            Dim col As Integer: col = 1
            Dim cellValue As Variant
            For Each cellValue In recapLineDetail
                If col = 4 Then
                    .Cells(row, col).value = cellValue & "%"
                Else
                    .Cells(row, col).value = cellValue
                    If col = 2 Then globalWork = globalWork + cellValue
                    If col = 3 Then globalActual = globalActual + cellValue
                End If
                col = col + 1
            Next
            row = row + 1
        Next
        
        ' Total global
        Dim globalPercent As Double
        If globalWork > 0 Then
            globalPercent = Round((globalActual / globalWork) * 100, 1)
        End If
        
        row = row + 1
        .Cells(row, 1).value = "TOTAL GÉNÉRAL"
        .Cells(row, 2).value = globalWork
        .Cells(row, 3).value = globalActual
        .Cells(row, 4).value = globalPercent & "%"
        
        ' Mise en forme récapitulatif
        .Range("A1:D1").Interior.Color = RGB(68, 114, 196)
        .Range("A1:D1").Font.Color = RGB(255, 255, 255)
        .Range("A1:D1").Font.Bold = True
        
        .Range("A" & row & ":D" & row).Font.Bold = True
        .Range("A" & row & ":D" & row).Interior.Color = RGB(217, 225, 242)
        .Columns.AutoFit
    End With

    Debug.Print "=== ONGLET 2: Écriture des données détaillées ==="
    
    ' === ONGLET 2 : DONNÉES DÉTAILLÉES ===
    Call WriteDetailSheet(xlDetailSheet, datesDesc, resList, totalPlanned, dailyActual, cumActual)
    Call FormatDetailSheet(xlDetailSheet)

    ' Sauvegarder et ouvrir
    xlBook.SaveAs fileName
    xlRecapSheet.Activate
    xlApp.Visible = True
    
    Dim dateCount As String
    If Not IsEmpty(datesAsc) And UBound(datesAsc) >= LBound(datesAsc) Then
        dateCount = CStr(UBound(datesAsc) + 1) & " dates réelles"
    Else
        dateCount = "0 dates réelles"
    End If
    
    MsgBox "? Export terminé :" & vbCrLf & _
           "?? Fichier Excel : " & fileName & vbCrLf & _
           "?? Onglet 1 : Récapitulatif (" & recapData.Count & " ressources)" & vbCrLf & _
           "?? Onglet 2 : Données détaillées (" & dateCount & ")" & vbCrLf & _
           "?? " & resList.Count & " ressource(s) mécanique(s) exportée(s)", vbInformation
    
    Shell "explorer.exe /select,""" & fileName & """", vbNormalFocus
    Debug.Print "=== FIN EXPORT MÉCANIQUE COMPLET ==="
    Exit Sub

ErrorHandler:
    Debug.Print "=== ERREUR DÉTECTÉE ==="
    Debug.Print "Erreur: " & Err.Description
    MsgBox "? Erreur lors de l'export : " & Err.Description, vbCritical
    
    On Error Resume Next
    If Not xlApp Is Nothing Then
        xlBook.Close False
        xlApp.Quit
    End If
End Sub

' === FONCTION DE TRI RAPIDE ===
Private Sub QuickSort(arr As Variant, first As Long, last As Long)
    Dim low As Long, high As Long, mid As String, temp As String
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
    If first < high Then QuickSort arr, first, high
    If low < last Then QuickSort arr, low, last
End Sub

' === NOUVELLES FONCTIONS REFACTORISÉES ===

' Tri des ressources mécaniques par ID de tâche ascendant
Private Function GetSortedMechanicalResources(proj As Project) As Collection
    Dim resInfo As Object
    Set resInfo = CreateObject("System.Collections.ArrayList")

    Dim res As Resource
    For Each res In proj.Resources
        If Not res Is Nothing Then
            Dim cleanGroup As String
            cleanGroup = Trim(Replace(Replace(res.Group, Chr(160), ""), Chr(32), " "))
            
            If (res.Type = 1 Or res.Type = 2) And _
               (cleanGroup = "Mécanique" Or UCase(cleanGroup) = "MÉCANIQUE") Then
                
                Dim minTaskId As Long
                minTaskId = 2147483647 ' Max value for Long
                
                If res.Assignments.Count > 0 Then
                    Dim assn As Assignment
                    For Each assn In res.Assignments
                        If assn.Task.ID < minTaskId Then
                            minTaskId = assn.Task.ID
                        End If
                    Next assn
                End If
                
                resInfo.Add Array(res.Name, minTaskId)
            End If
        End If
    Next res

    ' Tri
    If resInfo.Count > 1 Then
        Dim resArray As Variant
        resArray = resInfo.ToArray()
        QuickSortResources resArray, 0, resInfo.Count - 1
        resInfo.Clear
        
        Dim i As Long
        For i = 0 To UBound(resArray)
            resInfo.Add resArray(i)
        Next i
    End If
    
    ' Créer la collection de noms de ressources triée
    Dim sortedResList As Collection
    Set sortedResList = New Collection
    
    Dim item As Variant
    For Each item In resInfo
        sortedResList.Add item(0)
        Debug.Print "Ressource triée: " & item(0) & " (TaskID: " & item(1) & ")"
    Next
    
    Set GetSortedMechanicalResources = sortedResList
End Function

' Quicksort pour l'array de ressources (array de arrays)
Private Sub QuickSortResources(arr As Variant, first As Long, last As Long)
    Dim low As Long, high As Long
    Dim pivot As Long
    Dim temp As Variant
    
    low = first
    high = last
    pivot = arr((first + last) \ 2)(1) ' Pivot sur le minTaskId

    Do While low <= high
        ' Chercher un élément à gauche qui devrait être à droite
        Do While arr(low)(1) < pivot
            low = low + 1
        Loop
        ' Chercher un élément à droite qui devrait être à gauche
        Do While arr(high)(1) > pivot
            high = high - 1
        Loop
        
        If low <= high Then
            ' Échanger
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
    Set resAssignments = CreateObject("Scripting.Dictionary")
    
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
End Function

' Totaux prévus par ressource (Work)
Private Function ComputeTotalPlannedWork(resAssignments As Object) As Object
    Dim totalPlanned As Object
    Set totalPlanned = CreateObject("Scripting.Dictionary")
    
    Dim resName As Variant, assn As Assignment
    For Each resName In resAssignments.Keys
        Dim totalWork As Double: totalWork = 0
        For Each assn In resAssignments(resName)
            totalWork = totalWork + assn.Work
        Next
        totalPlanned(resName) = totalWork
    Next
    
    Set ComputeTotalPlannedWork = totalPlanned
End Function

' Travail réel par jour : Dict(resName -> Dict("yyyy-mm-dd" -> Double))
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
                        
                        If dailyActual(resName).exists(dateKey) Then
                            dailyActual(resName)(dateKey) = dailyActual(resName)(dateKey) + tsv(i).Value
                        Else
                            dailyActual(resName)(dateKey) = tsv(i).Value
                        End If
                    End If
                End If
            Next i
        Next
    Next
    
    Set ComputeDailyActualWork = dailyActual
End Function

' Dates où il y a du réel (union de toutes les ressources), triées
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
    Call QuickSort(sortedDates, LBound(sortedDates), UBound(sortedDates))
    
    If Not ascending Then
        sortedDates = ReverseArray(sortedDates)
    End If
    
    BuildActualDatesIndex = sortedDates
End Function

' Cumul par ressource et par date (dans le sens chronologique)
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

' Écriture onglet Données détaillées (Qté, Réel, Jour, %)
Private Sub WriteDetailSheet(xlWs As Object, orderedDatesDesc As Variant, _
    resOrder As Collection, totalPlanned As Object, dailyActual As Object, cumActual As Object)
    
    ' Déclarations de variables
    Dim col As Integer
    Dim resColMap As Object
    Dim resName As Variant
    Dim baseCol As Integer
    Dim row As Integer
    Dim d As Variant
    Dim realValue As Double
    Dim percentValue As Double
    
    ' En-têtes sur 2 lignes
    ' Ligne 1 : A1 vide, puis noms de ressources fusionnés sur 4 colonnes
    xlWs.Cells(1, 1).Value = ""
    
    col = 2
    Set resColMap = CreateObject("Scripting.Dictionary")
    
    For Each resName In resOrder
        ' Fusionner 4 cellules pour le nom de la ressource
        xlWs.Range(xlWs.Cells(1, col), xlWs.Cells(1, col + 3)).Merge
        xlWs.Cells(1, col).Value = resName
        xlWs.Cells(1, col).HorizontalAlignment = -4108  ' xlCenter
        
        resColMap(resName) = col
        col = col + 4
    Next
    
    ' Ligne 2 : "Date" en A2, puis sous-en-têtes pour chaque ressource
    xlWs.Cells(2, 1).Value = "Date"
    
    For Each resName In resOrder
        baseCol = resColMap(resName)
        xlWs.Cells(2, baseCol).Value = "Qté"
        xlWs.Cells(2, baseCol + 1).Value = "Réel"
        xlWs.Cells(2, baseCol + 2).Value = "Jour"
        xlWs.Cells(2, baseCol + 3).Value = "%"
    Next
    
    ' Données par date
    If IsEmpty(orderedDatesDesc) Or UBound(orderedDatesDesc) < LBound(orderedDatesDesc) Then
        xlWs.Cells(3, 1).Value = "Aucune donnée réelle trouvée"
        Exit Sub
    End If
    
    row = 3
    For Each d In orderedDatesDesc
        xlWs.Cells(row, 1).Value = Format(CDate(d), "dd/mm/yyyy")
        
        For Each resName In resOrder
            baseCol = resColMap(resName)
            
            ' Qté (total prévu)
            xlWs.Cells(row, baseCol).Value = totalPlanned(resName)
            
            ' Réel (cumul)
            realValue = 0
            If cumActual(resName).exists(d) Then
                realValue = cumActual(resName)(d)
                xlWs.Cells(row, baseCol + 1).Value = realValue
            Else
                xlWs.Cells(row, baseCol + 1).Value = 0
            End If
            
            ' Jour (travail du jour)
            If dailyActual(resName).exists(d) Then
                xlWs.Cells(row, baseCol + 2).Value = dailyActual(resName)(d)
            Else
                xlWs.Cells(row, baseCol + 2).Value = 0
            End If
            
            ' % (pourcentage Réel/Qté)
            percentValue = 0
            If totalPlanned(resName) > 0 Then
                percentValue = Round((realValue / totalPlanned(resName)) * 100, 1)
            End If
            xlWs.Cells(row, baseCol + 3).Value = percentValue & "%"
        Next
        row = row + 1
    Next
End Sub

' Mise en forme (fige ligne 1, formats, bordures)
Private Sub FormatDetailSheet(xlWs As Object)
    ' Figer la ligne 1
    xlWs.Range("A2").Select
    xlWs.Application.ActiveWindow.FreezePanes = True
    
    ' Mise en forme des en-têtes
    Dim lastCol As Integer: lastCol = xlWs.UsedRange.Columns.Count
    xlWs.Range("A1").Resize(1, lastCol).Interior.Color = RGB(68, 114, 196)
    xlWs.Range("A1").Resize(1, lastCol).Font.Color = RGB(255, 255, 255)
    xlWs.Range("A1").Resize(1, lastCol).Font.Bold = True
    
    ' Bordures fines
    Dim lastRow As Integer: lastRow = xlWs.UsedRange.Rows.Count
    If lastRow > 1 Then
        xlWs.Range("A1").Resize(lastRow, lastCol).Borders.LineStyle = 1  ' xlContinuous
        xlWs.Range("A1").Resize(lastRow, lastCol).Borders.Weight = 2     ' xlThin
    End If
    
    ' AutoFit
    xlWs.Columns.AutoFit
    
    ' === FORMAT D'AFFICHAGE SANS DÉCIMALES ===
    ' Appliquer le format entier/pourcentage à toutes les colonnes de données
    If lastRow > 2 Then ' S'il y a des données (au-delà des en-têtes)
        Dim formatCol As Integer
        
        ' Parcourir toutes les colonnes de données (colonnes 2 et suivantes, par groupes de 4)
        For formatCol = 2 To lastCol Step 4
            If formatCol <= lastCol Then
                ' Colonne "Qté" → format entier
                xlWs.Range(xlWs.Cells(3, formatCol), xlWs.Cells(lastRow, formatCol)).NumberFormat = "0"
            End If
            
            If formatCol + 1 <= lastCol Then
                ' Colonne "Réel" → format entier
                xlWs.Range(xlWs.Cells(3, formatCol + 1), xlWs.Cells(lastRow, formatCol + 1)).NumberFormat = "0"
            End If
            
            If formatCol + 2 <= lastCol Then
                ' Colonne "Jour" → format entier
                xlWs.Range(xlWs.Cells(3, formatCol + 2), xlWs.Cells(lastRow, formatCol + 2)).NumberFormat = "0"
            End If
            
            If formatCol + 3 <= lastCol Then
                ' Colonne "%" → format pourcentage entier (mais les valeurs sont déjà en % textuel)
                ' On garde le format texte car les valeurs contiennent déjà le symbole %
                ' Si on voulait un vrai format pourcentage : xlWs.Range(...).NumberFormat = "0%"
            End If
        Next formatCol
    End If
End Sub

' ===================================================================
' FONCTIONS POUR BOUTON ET VISIBILITÉ DANS "PERSONNALISER LE RUBAN"
' ===================================================================

' Bouton "Ruban" (via Personnaliser le ruban > Macros)
Public Sub ExportMeca_Bouton()
    ' Lance l'export complet (2 onglets, logique existante)
    ExportMecaniqueComplet
End Sub

' Crée un bouton dans l'onglet Compléments (Add-Ins) pour lancer l'export
Public Sub InstallerBoutonExportMeca()
    On Error Resume Next
    ' Nettoyage si déjà présent
    Application.CommandBars("ExportMeca").Delete
    On Error GoTo 0

    Dim cb As CommandBar
    Dim btn As CommandBarButton

    Set cb = Application.CommandBars.Add(Name:="ExportMeca", Position:=msoBarTop, Temporary:=True)
    Set btn = cb.Controls.Add(Type:=msoControlButton)

    With btn
        .Caption = "Export Mécanique"
        .OnAction = "ExportMeca_Bouton"  ' appelle le wrapper
        .Style = msoButtonIconAndCaption
        .FaceId = 176  ' icône standard Office ; changeable si besoin
        .TooltipText = "Exporter Récapitulatif + Données détaillées (Mécanique)"
    End With

    cb.Visible = True
End Sub

' Optionnel : suppression du bouton Compléments
Public Sub SupprimerBoutonExportMeca()
    On Error Resume Next
    Application.CommandBars("ExportMeca").Delete
    On Error GoTo 0
End Sub

