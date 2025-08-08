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
    Dim res As Resource, assn As Assignment
    Dim resName As Variant, dateKey As String
    Dim tsv As TimeScaleValues
    Dim xlApp As Object, xlBook As Object, xlRecapSheet As Object, xlDetailSheet As Object
    
    ' Collections pour les données
    Dim resList As Collection
    Dim resAssignments As Object
    Dim dataDict As Object  ' Pour les données détaillées (dates réelles)
    Dim dateDict As Object  ' Pour tracker les dates
    Dim recapData As Collection  ' Pour le récapitulatif

    Set resList = New Collection
    Set resAssignments = CreateObject("Scripting.Dictionary")
    Set dataDict = CreateObject("Scripting.Dictionary")
    Set dateDict = CreateObject("Scripting.Dictionary")
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

    Debug.Print "=== ÉTAPE 1: Collecte des ressources mécaniques ==="
    
    ' Collecter les ressources du groupe Mécanique
    For Each res In ActiveProject.Resources
        If Not res Is Nothing Then
            Dim cleanGroup As String
            cleanGroup = Trim(Replace(Replace(res.Group, Chr(160), ""), Chr(32), " "))
            
            If (res.Type = 1 Or res.Type = 2) And _
               (cleanGroup = "Mécanique" Or cleanGroup = "Mécanique" Or _
                UCase(cleanGroup) = "MÉCANIQUE" Or UCase(cleanGroup) = "MECANIQUE") Then
                
                Debug.Print "Ressource trouvée: " & res.Name
                resList.Add res.Name
                Set resAssignments(res.Name) = New Collection
                Set dataDict(res.Name) = CreateObject("Scripting.Dictionary")
            End If
        End If
    Next res

    If resList.Count = 0 Then
        MsgBox "?? Aucune ressource du groupe Mécanique trouvée dans le projet.", vbExclamation
        Exit Sub
    End If

    Debug.Print "=== ÉTAPE 2: Collecte des assignations ==="
    
    ' Collecter toutes les assignations des ressources mécaniques
    For Each res In ActiveProject.Resources
        If Not res Is Nothing And resAssignments.exists(res.Name) Then
            For Each assn In res.Assignments
                resAssignments(res.Name).Add assn
            Next
        End If
    Next

    Debug.Print "=== ÉTAPE 3: Calcul des données détaillées (dates réelles) ==="
    
    ' Collecter les données détaillées avec dates réelles uniquement
    For Each resName In resList
        For Each assn In resAssignments(resName)
            Set tsv = assn.TimeScaleData(startDate, endDate + 1, pjAssignmentTimescaledActualWork, pjTimescaleDays)
            Dim i As Integer
            For i = 1 To tsv.Count
                If Not tsv(i) Is Nothing And IsNumeric(tsv(i).value) Then
                    If tsv(i).value <> 0 Then
                        dateKey = Format(tsv(i).startDate, "yyyy-mm-dd")
                        
                        ' Accumuler les valeurs pour cette ressource/date
                        If dataDict(resName).exists(dateKey) Then
                            dataDict(resName)(dateKey) = dataDict(resName)(dateKey) + tsv(i).value
                        Else
                            dataDict(resName)(dateKey) = tsv(i).value
                        End If
                        
                        ' Tracker les dates uniques
                        If Not dateDict.exists(dateKey) Then dateDict.Add dateKey, True
                    End If
                End If
            Next i
        Next
    Next

    Debug.Print "=== ÉTAPE 4: Calcul du récapitulatif global ==="
    
    ' Calculer le récapitulatif global
    For Each resName In resList
        Dim recapTotalWork As Double: recapTotalWork = 0
        Dim recapTotalActual As Double: recapTotalActual = 0

        For Each assn In resAssignments(resName)
            recapTotalWork = recapTotalWork + assn.Work
            
            Set tsv = assn.TimeScaleData(startDate, endDate + 1, pjAssignmentTimescaledActualWork, pjTimescaleDays)
            For i = 1 To tsv.Count
                If Not tsv(i) Is Nothing And IsNumeric(tsv(i).value) Then
                    recapTotalActual = recapTotalActual + tsv(i).value
                End If
            Next i
        Next

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

    ' Tri des dates
    Dim sortedDates As Variant
    sortedDates = dateDict.Keys
    Call QuickSort(sortedDates, LBound(sortedDates), UBound(sortedDates))

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
    With xlDetailSheet
        ' En-têtes
        .Cells(1, 1).value = "Date"
        Dim c As Integer: c = 2
        Dim resColMap As Object: Set resColMap = CreateObject("Scripting.Dictionary")
        
        For Each resName In resList
            .Cells(1, c).value = resName
            resColMap(resName) = c
            c = c + 1
        Next
        
        ' Données par date (uniquement dates avec valeurs réelles)
        Dim r As Integer: r = 2
        Dim d As Variant
        For Each d In sortedDates
            .Cells(r, 1).value = Format(CDate(d), "dd/mm/yyyy")
            For Each resName In resList
                If dataDict(resName).exists(d) Then
                    .Cells(r, resColMap(resName)).value = Round(dataDict(resName)(d), 2)
                End If
            Next
            r = r + 1
        Next
        
        ' Mise en forme données détaillées
        .Range("A1").Resize(1, c - 1).Interior.Color = RGB(68, 114, 196)
        .Range("A1").Resize(1, c - 1).Font.Color = RGB(255, 255, 255)
        .Range("A1").Resize(1, c - 1).Font.Bold = True
        .Columns.AutoFit
    End With

    ' Sauvegarder et ouvrir
    xlBook.SaveAs fileName
    xlRecapSheet.Activate
    xlApp.Visible = True
    
    MsgBox "? Export terminé :" & vbCrLf & _
           "?? Fichier Excel : " & fileName & vbCrLf & _
           "?? Onglet 1 : Récapitulatif (" & recapData.Count & " ressources)" & vbCrLf & _
           "?? Onglet 2 : Données détaillées (" & UBound(sortedDates) + 1 & " dates réelles)" & vbCrLf & _
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

