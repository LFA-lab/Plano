' ===================================================================
' MACRO UNIFIEE : EXPORT MECANIQUE + ELECTRIQUE - 3 ONGLETS
' ===================================================================
' Export de toutes les ressources des groupes "Mecanique" et "Electrique"
' Critères: Groupe=Mecanique + Groupe=Electrique
' Onglet 1: Récapitulatif des ressources mécaniques ET électriques
' Onglet 2: Données détaillées mécaniques
' Onglet 3: Données détaillées électriques

' ===================================================================

' === FONCTION UTILITAIRE : SELECTEUR DE DOSSIER ===
Function PickFolder(ByVal defaultPath As String) As String
    ' Version simple compatible MS Project - utilise InputBox
    Dim userPath As String
    Dim promptMsg As String
    
    promptMsg = "Veuillez saisir le chemin du dossier d'export :" & vbCrLf & _
                "Exemple : C:\Users\VotreNom\Downloads" & vbCrLf & _
                "(ou laissez vide pour utiliser : " & defaultPath & ")"
    
    userPath = InputBox(promptMsg, "Dossier d'export", defaultPath)
    
    ' Si l'utilisateur annule (userPath = "") ou laisse vide, utiliser le défaut
    If userPath = "" Then
        PickFolder = defaultPath
        Exit Function
    End If
    
    ' Nettoyer le chemin (enlever les espaces, backslash final)
    userPath = Trim(userPath)
    If Right$(userPath, 1) = "\" Then
        userPath = Left$(userPath, Len(userPath) - 1)
    End If
    
    ' Vérifier que le dossier existe
    If Dir(userPath, vbDirectory) <> "" Then
        PickFolder = userPath
    Else
        MsgBox "Le dossier specifique n'existe pas :" & vbCrLf & userPath & vbCrLf & _
               "Utilisation du dossier par defaut : " & defaultPath, vbExclamation
        PickFolder = defaultPath
    End If
End Function

' === MACRO PRINCIPALE ===
Public Sub ExportMecanique()
    ' Lance directement l'export complet
    ExportMecaniqueComplet
End Sub

Sub ExportMecaniqueComplet()
    Dim currentStep As String
    currentStep = "INITIALISATION"
    
    Debug.Print "=== DEBUT EXPORT MECANIQUE COMPLET === " & Format(Now, "hh:nn:ss")
    
    Dim fileName As String, exportDir As String
    Dim startDate As Date, endDate As Date
    Dim xlApp As Object, xlBook As Object, xlRecapSheet As Object, xlDetailSheet As Object
    
    ' Collections pour les donnees
    Dim resList As Collection
    Dim resAssignments As Object
    Dim totalPlanned As Object
    Dim dailyActual As Object
    Dim cumActual As Object
    Dim datesAsc As Variant, datesDesc As Variant
    Dim recapData As Collection

    Set recapData = New Collection

    On Error GoTo ErrorHandler

    ' Choisir le dossier d'export avec vérification
    currentStep = "SELECTION_DOSSIER"
    Debug.Print "Start: " & currentStep & " | " & Format(Now, "hh:nn:ss")
    
    Dim defaultDir As String
    defaultDir = Environ$("USERPROFILE") & "\Downloads"
    
    ' Vérifier que le dossier par défaut existe, sinon utiliser le Bureau
    If Dir(defaultDir, vbDirectory) = "" Then
        defaultDir = Environ$("USERPROFILE") & "\Desktop"
        If Dir(defaultDir, vbDirectory) = "" Then
            defaultDir = Environ$("USERPROFILE") & "\Documents"
        End If
    End If
    
    exportDir = PickFolder(defaultDir)
    Debug.Print "Dossier selectionne: " & exportDir
    
    ' Vérifier le retour de PickFolder
    If exportDir = "" Then
        MsgBox "Aucun dossier valide sélectionné. Export annulé.", vbInformation
        Exit Sub
    End If

    If Right$(exportDir, 1) = "\" Then
        exportDir = Left$(exportDir, Len(exportDir) - 1)
    End If

    fileName = exportDir & "\Export_Mecanique_" & Format(Now, "yyyymmdd_hhnnss") & ".xlsx"
    Debug.Print "Nom fichier: " & fileName
    
    startDate = ActiveProject.ProjectStart
    endDate = ActiveProject.ProjectFinish
    Debug.Print "Done: " & currentStep & " | " & Format(Now, "hh:nn:ss")

    currentStep = "COLLECTE_RESSOURCES"
    Debug.Print "Start: " & currentStep & " | " & Format(Now, "hh:nn:ss")
    Debug.Print "=== ETAPE 1: Collecte et tri des ressources mecaniques ==="
    
    ' Collecter les ressources pour chaque feuille
    Set resList = GetSortedMechanicalResources(ActiveProject, "") ' Pour le recapitulatif global (toutes ressources mecaniques)
    Dim resListMecanique As Collection, resListElectrique As Collection, resListGlobal As Collection
    Set resListMecanique = GetSortedMechanicalResources(ActiveProject, "") ' Pour la feuille détaillée mécanique
    Set resListElectrique = GetSortedElectricalResources(ActiveProject)
    
    ' Créer une collection globale pour le récapitulatif (mécaniques + électriques)
    Set resListGlobal = New Collection
    Dim resName As Variant
    For Each resName In resListMecanique
        resListGlobal.Add resName
    Next resName
    For Each resName In resListElectrique
        resListGlobal.Add resName
    Next resName
    
    Debug.Print "Nombre de ressources trouvees - Global: " & resListGlobal.Count & ", Mecanique: " & resListMecanique.Count & ", Electrique: " & resListElectrique.Count
    
    If resListMecanique.Count = 0 Then
        MsgBox "Aucune ressource du groupe Mecanique trouvee dans le projet.", vbExclamation
        Exit Sub
    End If
    
    If resListElectrique.Count = 0 Then
        MsgBox "Aucune ressource du groupe Electrique trouvee dans le projet.", vbExclamation
        Exit Sub
    End If
    Debug.Print "Done: " & currentStep & " | " & Format(Now, "hh:nn:ss")

    currentStep = "COLLECTE_ASSIGNATIONS"
    Debug.Print "Start: " & currentStep & " | " & Format(Now, "hh:nn:ss")
    Debug.Print "=== ETAPE 2: Collecte des assignations ==="
    
    ' Assignations pour recapitulatif (toutes les ressources: mecaniques + electriques)
    Set resAssignments = MapAssignmentsByResource(resListGlobal)
    
    ' Assignations specifiques pour Mecanique et Electrique
    Dim resAssignmentsMecanique As Object, resAssignmentsElectrique As Object
    Set resAssignmentsMecanique = MapAssignmentsByResource(resListMecanique)
    Set resAssignmentsElectrique = MapAssignmentsByResource(resListElectrique)
    Debug.Print "Done: " & currentStep & " | " & Format(Now, "hh:nn:ss")

    currentStep = "CALCUL_DONNEES"
    Debug.Print "Start: " & currentStep & " | " & Format(Now, "hh:nn:ss")
    Debug.Print "=== ETAPE 3: Calcul des donnees ==="
    
    ' Calculs pour recapitulatif global
    Set totalPlanned = ComputeTotalPlannedWork(resAssignments)
    Set dailyActual = ComputeDailyActualWork(resAssignments, startDate, endDate)
    datesAsc = BuildActualDatesIndex(dailyActual, True)
    Set cumActual = ComputeCumulativeActual(dailyActual, datesAsc)
    datesDesc = ReverseArray(datesAsc)
    Debug.Print "Calculs globaux termines"
    
    ' Calculs pour feuille Mecanique
    Dim totalPlannedMecanique As Object, dailyActualMecanique As Object, cumActualMecanique As Object
    Dim datesAscMecanique As Variant, datesDescMecanique As Variant
    Set totalPlannedMecanique = ComputeTotalPlannedWork(resAssignmentsMecanique)
    Set dailyActualMecanique = ComputeDailyActualWork(resAssignmentsMecanique, startDate, endDate)
    datesAscMecanique = BuildActualDatesIndex(dailyActualMecanique, True)
    Set cumActualMecanique = ComputeCumulativeActual(dailyActualMecanique, datesAscMecanique)
    datesDescMecanique = ReverseArray(datesAscMecanique)
    Debug.Print "Calculs Mecanique termines"
    
    ' Calculs pour feuille Electrique
    Dim totalPlannedElectrique As Object, dailyActualElectrique As Object, cumActualElectrique As Object
    Dim datesAscElectrique As Variant, datesDescElectrique As Variant
    Set totalPlannedElectrique = ComputeTotalPlannedWork(resAssignmentsElectrique)
    Set dailyActualElectrique = ComputeDailyActualWork(resAssignmentsElectrique, startDate, endDate)
    datesAscElectrique = BuildActualDatesIndex(dailyActualElectrique, True)
    Set cumActualElectrique = ComputeCumulativeActual(dailyActualElectrique, datesAscElectrique)
    datesDescElectrique = ReverseArray(datesAscElectrique)
    Debug.Print "Calculs Electrique termines"
    Debug.Print "Done: " & currentStep & " | " & Format(Now, "hh:nn:ss")

    currentStep = "CALCUL_RECAPITULATIF"
    Debug.Print "Start: " & currentStep & " | " & Format(Now, "hh:nn:ss")
    Debug.Print "=== ETAPE 4: Calcul du recapitulatif global par ressource consolidee ==="
    
    ' Créer un dictionnaire pour consolider les ressources par nom (mécaniques + électriques)
    Dim consolidatedResources As Object
    Set consolidatedResources = CreateObject("Scripting.Dictionary")
    
    ' Fonction pour nettoyer le nom de ressource (pas de suffixes spéciaux à enlever)
    For Each resName In resListGlobal
        Dim cleanName As String
        cleanName = CStr(resName)
        
        ' Nettoyer le nom de ressource (pas de suffixes spéciaux à enlever)
        cleanName = Trim(cleanName)
        
        ' Ignorer les ressources Implantation
        If UCase(cleanName) = "IMPLANTATION" Then
            Debug.Print "Ressource Implantation ignor�e: " & resName
            GoTo NextResName
        End If
        
        ' Calculer les valeurs pour cette ressource
        Dim recapTotalWork As Double: recapTotalWork = 0
        Dim recapTotalActual As Double: recapTotalActual = 0
        
        ' V�rifier si totalPlanned contient cette ressource
        If totalPlanned.exists(resName) Then
            recapTotalWork = totalPlanned(resName)
        End If
        
        ' Calculer le total reel (maximum du cumul)
        If Not IsEmpty(datesAsc) And UBound(datesAsc) >= LBound(datesAsc) Then
            Dim lastDate As String: lastDate = datesAsc(UBound(datesAsc))
            If cumActual.exists(resName) And cumActual(resName).exists(lastDate) Then
                recapTotalActual = cumActual(resName)(lastDate)
            End If
        End If
        
        ' Ajouter ou consolider dans le dictionnaire
        If consolidatedResources.exists(cleanName) Then
            ' Ajouter aux valeurs existantes
            Dim existingValues As Variant
            existingValues = consolidatedResources(cleanName)
            ' existingValues est un array: [0]=Pr�vu, [1]=R�alis�
            existingValues(0) = existingValues(0) + recapTotalWork
            existingValues(1) = existingValues(1) + recapTotalActual
            consolidatedResources(cleanName) = existingValues
        Else
            ' Cr�er nouvelle entr�e
            Dim newValues(1) As Double
            newValues(0) = recapTotalWork   ' Pr�vu
            newValues(1) = recapTotalActual ' R�alis�
            consolidatedResources.Add cleanName, newValues
        End If
        
        Debug.Print "Ressource " & resName & " -> " & cleanName & ": " & recapTotalWork & "/" & recapTotalActual
NextResName:
    Next
    
    ' Cr�er les lignes du r�capitulatif � partir des ressources consolid�es
    Dim consolidatedName As Variant
    For Each consolidatedName In consolidatedResources.Keys
        Dim consolidatedValues As Variant
        consolidatedValues = consolidatedResources(consolidatedName)
        
        Dim totalWork As Double: totalWork = consolidatedValues(0)
        Dim totalActual As Double: totalActual = consolidatedValues(1)
        
        Dim recapPercent As Double
        If totalWork > 0 Then
            recapPercent = Round((totalActual / totalWork) * 100, 1)
        End If
        
        Dim recapLineCalc As Collection
        Set recapLineCalc = New Collection
        recapLineCalc.Add consolidatedName
        recapLineCalc.Add Round(totalWork, 0)
        recapLineCalc.Add Round(totalActual, 0)
        recapLineCalc.Add recapPercent
        recapData.Add recapLineCalc
        
        Debug.Print "Recap consolid� " & consolidatedName & ": " & totalWork & "/" & totalActual & " (" & recapPercent & "%)"
    Next
    Debug.Print "Done: " & currentStep & " | " & Format(Now, "hh:nn:ss")

    currentStep = "CREATION_EXCEL"
    Debug.Print "Start: " & currentStep & " | " & Format(Now, "hh:nn:ss")
    Debug.Print "=== ETAPE 5: Creation du fichier Excel ==="

    ' Creer Excel avec 2 onglets - Gestion d'erreur amelioree
    On Error GoTo ExcelError
    Set xlApp = CreateObject("Excel.Application")
    If xlApp Is Nothing Then
        MsgBox "Erreur : Impossible de creer l'application Excel." & vbCrLf & _
               "Verifiez qu'Excel est installe et accessible.", vbCritical
        Exit Sub
    End If
    Debug.Print "Application Excel créée"
    
    Set xlBook = xlApp.Workbooks.Add
    Debug.Print "Classeur Excel créé"
    On Error GoTo ErrorHandler
    
    ' Supprimer feuilles par defaut sauf une
    xlApp.DisplayAlerts = False
    Do While xlBook.Worksheets.Count > 1
        xlBook.Worksheets(xlBook.Worksheets.Count).Delete
    Loop
    xlApp.DisplayAlerts = True
    Debug.Print "Feuilles par defaut supprimees"
    
    ' Creer les 3 onglets
    Set xlRecapSheet = xlBook.Worksheets(1)
    xlRecapSheet.Name = "Recapitulatif"
    Debug.Print "Onglet Recapitulatif cree"
    
    Dim xlDetailSheetMecanique As Object, xlDetailSheetElectrique As Object
    Set xlDetailSheetMecanique = xlBook.Worksheets.Add
    xlDetailSheetMecanique.Name = "Donnees Mecanique"
    xlDetailSheetMecanique.Move After:=xlRecapSheet
    Debug.Print "Onglet Donnees Mecanique cree"
    
    Set xlDetailSheetElectrique = xlBook.Worksheets.Add
    xlDetailSheetElectrique.Name = "Donnees Electrique"
    xlDetailSheetElectrique.Move After:=xlDetailSheetMecanique
    Debug.Print "Onglet Donnees Electrique cree"
    Debug.Print "Done: " & currentStep & " | " & Format(Now, "hh:nn:ss")

    currentStep = "INSERTION_LOGOS"
    Debug.Print "Start: " & currentStep & " | " & Format(Now, "hh:nn:ss")
    Debug.Print "=== INSERTION DES LOGOS ==="
    Call InsererLogos(xlRecapSheet)
    Debug.Print "Done: " & currentStep & " | " & Format(Now, "hh:nn:ss")

    currentStep = "ONGLET_RECAPITULATIF"
    Debug.Print "Start: " & currentStep & " | " & Format(Now, "hh:nn:ss")
    Debug.Print "=== ONGLET 1: Ecriture du recapitulatif ==="
    
    ' === ONGLET 1 : RECAPITULATIF ===
    With xlRecapSheet
        ' Décaler les données vers le bas pour laisser place aux logos
        .Cells(4, 1).Value = "Ressource"
        .Cells(4, 2).Value = "Prevu"
        .Cells(4, 3).Value = "Realise"
        .Cells(4, 4).Value = "Pourcentage"
        
        Dim row As Integer: row = 5
        Dim recapLineDetail As Collection
        Dim globalWork As Double: globalWork = 0
        Dim globalActual As Double: globalActual = 0
        
        For Each recapLineDetail In recapData
            Dim col As Integer: col = 1
            Dim cellValue As Variant
            
            For Each cellValue In recapLineDetail
                If col = 4 Then
                    .Cells(row, col).Value = cellValue & "%"
                Else
                    .Cells(row, col).Value = cellValue
                    If col = 2 And IsNumeric(cellValue) Then globalWork = globalWork + cellValue
                    If col = 3 And IsNumeric(cellValue) Then globalActual = globalActual + cellValue
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
        .Cells(row, 1).Value = "TOTAL GENERAL"
        .Cells(row, 2).Value = globalWork
        .Cells(row, 3).Value = globalActual
        .Cells(row, 4).Value = globalPercent & "%"
        
        ' Mise en forme recapitulatif
        .Range("A4:D4").Interior.Color = RGB(68, 114, 196)
        .Range("A4:D4").Font.Color = RGB(255, 255, 255)
        .Range("A4:D4").Font.Bold = True
        
        .Range("A" & row & ":D" & row).Font.Bold = True
        .Range("A" & row & ":D" & row).Interior.Color = RGB(217, 225, 242)
        .Columns.AutoFit
    End With
    Debug.Print "Done: " & currentStep & " | " & Format(Now, "hh:nn:ss")

    currentStep = "ONGLET_DETAILS"
    Debug.Print "Start: " & currentStep & " | " & Format(Now, "hh:nn:ss")
    Debug.Print "=== ONGLETS 2 et 3: Ecriture des donnees detaillees ==="
    
    ' === ONGLET 2 : DONNEES MECANIQUE ===
    If resListMecanique.Count > 0 Then
        Call WriteDetailSheet(xlDetailSheetMecanique, datesDescMecanique, resListMecanique, totalPlannedMecanique, dailyActualMecanique, cumActualMecanique)
        Debug.Print "Données Mecanique écrites"
        Call FormatDetailSheet(xlDetailSheetMecanique)
        Debug.Print "Formatage Mecanique terminé"
    Else
        xlDetailSheetMecanique.Cells(1, 1).Value = "Aucune ressource du groupe Mecanique trouvee"
        Debug.Print "Aucune donnée Mecanique à écrire"
    End If
    
    ' === ONGLET 3 : DONNEES ELECTRIQUE ===
    If resListElectrique.Count > 0 Then
        Call WriteDetailSheet(xlDetailSheetElectrique, datesDescElectrique, resListElectrique, totalPlannedElectrique, dailyActualElectrique, cumActualElectrique)
        Debug.Print "Données Electrique écrites"
        Call FormatDetailSheet(xlDetailSheetElectrique)
        Debug.Print "Formatage Electrique terminé"
    Else
        xlDetailSheetElectrique.Cells(1, 1).Value = "Aucune ressource du groupe Electrique trouvee"
        Debug.Print "Aucune donnée Electrique à écrire"
    End If
    Debug.Print "Done: " & currentStep & " | " & Format(Now, "hh:nn:ss")

    ' Sauvegarder et ouvrir
    currentStep = "SAUVEGARDE"
    Debug.Print "Start: " & currentStep & " | " & Format(Now, "hh:nn:ss")
    Debug.Print "Fichier de sauvegarde: " & fileName
    
    xlBook.SaveAs fileName
    Debug.Print "Fichier sauvegardé avec succès"
    
    xlRecapSheet.Activate
    xlApp.Visible = True
    Debug.Print "Done: " & currentStep & " | " & Format(Now, "hh:nn:ss")
    
    currentStep = "FINALISATION"
    Debug.Print "Start: " & currentStep & " | " & Format(Now, "hh:nn:ss")
    
    Dim dateCount As String
    If Not IsEmpty(datesAsc) And UBound(datesAsc) >= LBound(datesAsc) Then
        dateCount = CStr(UBound(datesAsc) + 1) & " dates reelles"
    Else
        dateCount = "0 dates reelles"
    End If
    
    MsgBox "Export termine :" & vbCrLf & _
           "Fichier Excel : " & fileName & vbCrLf & _
           "Onglet 1 : Recapitulatif consolid� (" & recapData.Count & " types de ressources)" & vbCrLf & _
           "Onglet 2 : Donnees Mecanique (" & resListMecanique.Count & " ressources)" & vbCrLf & _
           "Onglet 3 : Donnees Electrique (" & resListElectrique.Count & " ressources)" & vbCrLf & _
           "Total : " & resListGlobal.Count & " ressource(s) au total", vbInformation
    
    ' Tentative d'ouvrir l'explorateur (peut echouer selon les politiques de securite)
    On Error Resume Next
    Shell "explorer.exe /select,""" & fileName & """", vbNormalFocus
    If Err.Number <> 0 Then
        Debug.Print "Info : Ouverture automatique de l'explorateur non autorisee"
    End If
    On Error GoTo 0
    Debug.Print "Done: " & currentStep & " | " & Format(Now, "hh:nn:ss")
    Debug.Print "=== FIN EXPORT MECANIQUE COMPLET === " & Format(Now, "hh:nn:ss")
    Exit Sub

ErrorHandler:
    Debug.Print "=== ERREUR DETECTEE à l'étape: " & currentStep & " | " & Format(Now, "hh:nn:ss") & " ==="
    Debug.Print "ERREUR: Err=" & Err.Number & " - " & Err.Description
    
    MsgBox "Erreur lors de l'export à l'étape: " & currentStep & vbCrLf & _
           "Erreur: " & Err.Number & " - " & Err.Description & vbCrLf & _
           "Consultez la fenêtre Immediate (Ctrl+G) pour plus de détails.", vbCritical
    
    On Error Resume Next
    If Not xlApp Is Nothing Then
        If Not xlBook Is Nothing Then xlBook.Close False
        xlApp.Quit
        Debug.Print "Application Excel fermée en urgence"
    End If
    On Error GoTo 0
    Exit Sub

ExcelError:
    Debug.Print "=== ERREUR EXCEL à l'étape: " & currentStep & " | " & Format(Now, "hh:nn:ss") & " ==="
    Debug.Print "ERREUR EXCEL: Err=" & Err.Number & " - " & Err.Description
    
    MsgBox "Probleme d'automation Excel detecte à l'étape: " & currentStep & vbCrLf & _
           "- Verifiez qu'Excel est installe" & vbCrLf & _
           "- Fermez Excel s'il est ouvert" & vbCrLf & _
           "- Redemarrez MS Project" & vbCrLf & _
           "Erreur: " & Err.Number & " - " & Err.Description, vbCritical
    Exit Sub
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

' === NOUVELLES FONCTIONS REFACTORISEES ===

' Tri des ressources mecaniques par ID de tache ascendant
Private Function GetSortedMechanicalResources(proj As Project, Optional zoneFilter As String = "") As Collection
    Debug.Print "Start: Collecte ressources mecaniques specifiques" & IIf(zoneFilter <> "", " (Zone=" & zoneFilter & ")", "") & " | " & Format(Now, "hh:nn:ss")
    
    On Error GoTo ErrorHandler

    ' Utiliser le groupe "Mecanique" au lieu d'une liste hardcodée

    
    ' Compteurs pour le logging (plus besoin de Est/Ouest)
    Dim countTotal As Long
    countTotal = 0
    
    ' Utiliser un tableau dynamique
    Dim resArray() As Variant
    Dim resCount As Long
    resCount = 0
    
    ' Redimensionner le tableau initial
    ReDim resArray(0 To 50) ' Taille initiale pour les ressources du groupe Mecanique

    Dim res As Resource
    For Each res In proj.Resources
        If Not res Is Nothing Then
            Dim resName As String
            resName = Trim(res.Name)
            
            ' Vérifier si la ressource appartient au groupe "Mecanique"
            If UCase(res.Group) <> "MECANIQUE" Then
                GoTo NextResource
            End If
            
            ' Appliquer le filtre de zone si sp�cifi�
            Dim zoneMatch As Boolean
            If zoneFilter = "" Then
                ' Pour le r�capitulatif global : accepter toutes les ressources valides
                zoneMatch = True
                countTotal = countTotal + 1
                
                ' Comptage pour le log (plus de distinction Est/Ouest)
            Else
                ' Pour les feuilles specifiques : filtrer selon la zone demandée
                zoneMatch = (InStr(1, resName, " " & zoneFilter, vbTextCompare) > 0)
                If zoneMatch Then countTotal = countTotal + 1
            End If
            
            If zoneMatch Then
                
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
                
                ' Redimensionner le tableau si nécessaire
                If resCount > UBound(resArray) Then
                    ReDim Preserve resArray(0 To UBound(resArray) + 50)
                End If
                
                resArray(resCount) = Array(res.Name, minTaskId)
                resCount = resCount + 1
            End If
            
NextResource:
        End If
    Next res
    
    ' Log d�taill� des ressources trouv�es
    If zoneFilter = "" Then
        Debug.Print "Ressources valides trouvees: " & resCount
    Else
        Debug.Print "Ressources valides trouvees (Zone " & zoneFilter & "): " & resCount
    End If

    ' Redimensionner le tableau à la taille exacte
    If resCount > 0 Then
        ReDim Preserve resArray(0 To resCount - 1)
        
        ' Tri si on a plus d'une ressource
        If resCount > 1 Then
            QuickSortResources resArray, 0, resCount - 1
        End If
    End If
    
    ' Creer la collection de noms de ressources triee
    Dim sortedResList As Collection
    Set sortedResList = New Collection
    
    Dim i As Long
    For i = 0 To resCount - 1
        sortedResList.Add resArray(i)(0)
        Debug.Print "Ressource triee: " & resArray(i)(0) & " (TaskID: " & resArray(i)(1) & ")"
    Next i
    
    Debug.Print "Done: Collecte ressources mecaniques | " & Format(Now, "hh:nn:ss")
    Set GetSortedMechanicalResources = sortedResList
    Exit Function
    
ErrorHandler:
    Debug.Print "Erreur dans GetSortedMechanicalResources: " & Err.Description
    Set sortedResList = New Collection
    Set GetSortedMechanicalResources = sortedResList
End Function

' Fonction pour collecter les ressources du groupe Electrique
Private Function GetSortedElectricalResources(proj As Project) As Collection
    Debug.Print "Start: Collecte ressources electriques | " & Format(Now, "hh:nn:ss")
    
    On Error GoTo ErrorHandler

    ' Utiliser un tableau dynamique
    Dim resArray() As Variant
    Dim resCount As Long
    resCount = 0
    
    ' Redimensionner le tableau initial
    ReDim resArray(0 To 50) ' Taille initiale pour les ressources du groupe Electrique

    Dim res As Resource
    For Each res In proj.Resources
        If Not res Is Nothing Then
            Dim resName As String
            resName = Trim(res.Name)
            
            ' Vérifier si la ressource appartient au groupe "Electrique"
            If UCase(res.Group) <> "ELECTRIQUE" Then
                GoTo NextResource
            End If
            
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
            
            ' Redimensionner le tableau si nécessaire
            If resCount > UBound(resArray) Then
                ReDim Preserve resArray(0 To UBound(resArray) + 50)
            End If
            
            resArray(resCount) = Array(res.Name, minTaskId)
            resCount = resCount + 1
            
NextResource:
        End If
    Next res
    
    Debug.Print "Ressources electriques trouvees: " & resCount

    ' Redimensionner le tableau à la taille exacte
    If resCount > 0 Then
        ReDim Preserve resArray(0 To resCount - 1)
        
        ' Tri si on a plus d'une ressource
        If resCount > 1 Then
            QuickSortResources resArray, 0, resCount - 1
        End If
    End If
    
    ' Creer la collection de noms de ressources triee
    Dim sortedResList As Collection
    Set sortedResList = New Collection
    
    Dim i As Long
    For i = 0 To resCount - 1
        sortedResList.Add resArray(i)(0)
        Debug.Print "Ressource electrique triee: " & resArray(i)(0) & " (TaskID: " & resArray(i)(1) & ")"
    Next i
    
    Debug.Print "Done: Collecte ressources electriques | " & Format(Now, "hh:nn:ss")
    Set GetSortedElectricalResources = sortedResList
    Exit Function
    
ErrorHandler:
    Debug.Print "Erreur dans GetSortedElectricalResources: " & Err.Description
    Set sortedResList = New Collection
    Set GetSortedElectricalResources = sortedResList
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
        ' Chercher un Ã©lÃ©ment Ã  gauche qui devrait Ãªtre Ã  droite
        Do While arr(low)(1) < pivot
            low = low + 1
        Loop
        ' Chercher un Ã©lÃ©ment Ã  droite qui devrait Ãªtre Ã  gauche
        Do While arr(high)(1) > pivot
            high = high - 1
        Loop
        
        If low <= high Then
            ' Ã‰changer
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
    Debug.Print "Start: Index assignations par ressource | " & Format(Now, "hh:nn:ss")
    
    Dim resAssignments As Object
    
    ' Tentative de creation d'un Dictionary avec gestion d'erreur
    On Error GoTo DictError
    Set resAssignments = CreateObject("Scripting.Dictionary")
    On Error GoTo 0
    
    Dim resName As Variant
    For Each resName In resList
        Set resAssignments(resName) = New Collection
    Next
    
    Dim res As Resource, assn As Assignment
    Dim totalAssignments As Long: totalAssignments = 0
    
    For Each res In ActiveProject.Resources
        If Not res Is Nothing And resAssignments.exists(res.Name) Then
            For Each assn In res.Assignments
                resAssignments(res.Name).Add assn
                totalAssignments = totalAssignments + 1
            Next
        End If
    Next
    
    Debug.Print "Assignations trouvees: " & totalAssignments
    Debug.Print "Done: Index assignations par ressource | " & Format(Now, "hh:nn:ss")
    Set MapAssignmentsByResource = resAssignments
    Exit Function
    
DictError:
    Debug.Print "ERREUR creation Dictionary: " & Err.Number & " - " & Err.Description
    MsgBox "Erreur : Impossible de creer l'objet Dictionary." & vbCrLf & _
           "Verifiez que Microsoft Scripting Runtime est disponible.", vbCritical
End Function

' Totaux prevus par ressource (Work)
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

' Travail reel par jour : Dict(resName -> Dict("yyyy-mm-dd" -> Double))
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

' Dates ou il y a du reel (union de toutes les ressources), triees
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

' Ecriture onglet Donnees detaillees (Qte, Reel, Jour, %)
Private Sub WriteDetailSheet(xlWs As Object, orderedDatesDesc As Variant, _
    resOrder As Collection, totalPlanned As Object, dailyActual As Object, cumActual As Object)

    Debug.Print "Start: Ecriture onglet detaille | " & Format(Now, "hh:nn:ss")

    ' Declarations de variables
    Dim col As Integer
    Dim resColMap As Object
    Dim resName As Variant
    Dim baseCol As Integer
    Dim row As Integer
    Dim d As Variant
    Dim realValue As Double
    Dim percentValue As Double
    
    ' En-tetes sur 2 lignes
    ' Ligne 1 : A1 vide, puis noms de ressources fusionnes sur 4 colonnes
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
    Debug.Print "En-tetes ressources cree"

    ' Ligne 2 : "Date" en A2, puis sous-en-tetes pour chaque ressource
    xlWs.Cells(2, 1).Value = "Date"
    
    For Each resName In resOrder
        baseCol = resColMap(resName)
        xlWs.Cells(2, baseCol).Value = "Qte"
        xlWs.Cells(2, baseCol + 1).Value = "Reel"
        xlWs.Cells(2, baseCol + 2).Value = "Jour"
        xlWs.Cells(2, baseCol + 3).Value = "%"
    Next
    Debug.Print "Sous-en-tetes cree"

    ' Donnees par date
    If IsEmpty(orderedDatesDesc) Or UBound(orderedDatesDesc) < LBound(orderedDatesDesc) Then
        xlWs.Cells(3, 1).Value = "Aucune donnee reelle trouvee"
        Debug.Print "Aucune donnee reelle trouvee"
        Exit Sub
    End If
    
    row = 3
    Dim dateCount As Long: dateCount = 0
    For Each d In orderedDatesDesc
        xlWs.Cells(row, 1).Value = Format(CDate(d), "dd/mm/yyyy")
        
        For Each resName In resOrder
            baseCol = resColMap(resName)
            
            ' Qte (total prevu)
            xlWs.Cells(row, baseCol).Value = totalPlanned(resName)
            
            ' Reel (cumul)
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
            
            ' % (pourcentage Reel/Qte)
            percentValue = 0
            If totalPlanned(resName) > 0 Then
                percentValue = Round((realValue / totalPlanned(resName)) * 100, 1)
            End If
            xlWs.Cells(row, baseCol + 3).Value = percentValue & "%"
        Next
        row = row + 1
        dateCount = dateCount + 1
    Next

    Debug.Print "Donnees ecrites: " & dateCount & " dates"
    Debug.Print "Done: Ecriture onglet detaille | " & Format(Now, "hh:nn:ss")
End Sub

' Mise en forme (fige ligne 1, formats, bordures)
Private Sub FormatDetailSheet(xlWs As Object)
    On Error Resume Next
    ' Figer la ligne 1 - seulement si Excel est visible
    If xlWs.Application.Visible Then
        xlWs.Activate
        xlWs.Range("A2").Select
        xlWs.Application.ActiveWindow.FreezePanes = True
    End If
    On Error GoTo 0
    
    ' Mise en forme des en-tÃªtes
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
    
    ' === FORMAT D'AFFICHAGE SANS DÃ‰CIMALES ===
    ' Appliquer le format entier/pourcentage Ã  toutes les colonnes de donnÃ©es
    If lastRow > 2 Then ' S'il y a des donnÃ©es (au-delÃ  des en-tÃªtes)
        Dim formatCol As Integer
        Dim colorIndex As Integer
        
        ' Palette de couleurs plus fonc�es pour les ressources consommables
        Dim resourceColors As Variant
        resourceColors = Array(RGB(255, 182, 193), RGB(144, 238, 144), RGB(173, 216, 230), _
                              RGB(255, 255, 0), RGB(221, 160, 221), RGB(0, 255, 255), _
                              RGB(255, 165, 0), RGB(154, 205, 50), RGB(135, 206, 250), _
                              RGB(255, 105, 180), RGB(186, 85, 211), RGB(0, 250, 154))
        
        colorIndex = 0
        
        ' Parcourir toutes les colonnes de donnÃ©es (colonnes 2 et suivantes, par groupes de 4)
        For formatCol = 2 To lastCol Step 4
            ' Appliquer la couleur de fond pour le groupe de 4 colonnes de la ressource
            Dim currentColor As Long
            currentColor = resourceColors(colorIndex Mod UBound(resourceColors) + 1)
            
            ' Appliquer la couleur aux en-t�tes (lignes 1 et 2) et aux donn�es pour les 4 colonnes
            If formatCol + 3 <= lastCol Then
                ' Colorer les en-t�tes de la ressource (ligne 1 fusionn�e et ligne 2)
                xlWs.Range(xlWs.Cells(1, formatCol), xlWs.Cells(2, formatCol + 3)).Interior.Color = currentColor
                
                ' Mettre le texte des noms de ressources en noir (ligne 1)
                xlWs.Range(xlWs.Cells(1, formatCol), xlWs.Cells(1, formatCol + 3)).Font.Color = RGB(0, 0, 0)
                
                ' Mettre le texte des sous-en-t�tes en noir (ligne 2)
                xlWs.Range(xlWs.Cells(2, formatCol), xlWs.Cells(2, formatCol + 3)).Font.Color = RGB(0, 0, 0)
                
                ' Colorer toutes les donn�es de cette ressource
                xlWs.Range(xlWs.Cells(3, formatCol), xlWs.Cells(lastRow, formatCol + 3)).Interior.Color = currentColor
            End If
            If formatCol <= lastCol Then
                ' Colonne "QtÃ©" â†’ format entier
                xlWs.Range(xlWs.Cells(3, formatCol), xlWs.Cells(lastRow, formatCol)).NumberFormat = "0"
            End If
            
            If formatCol + 1 <= lastCol Then
                ' Colonne "RÃ©el" â†’ format entier
                xlWs.Range(xlWs.Cells(3, formatCol + 1), xlWs.Cells(lastRow, formatCol + 1)).NumberFormat = "0"
            End If
            
            If formatCol + 2 <= lastCol Then
                ' Colonne "Jour" â†’ format entier
                xlWs.Range(xlWs.Cells(3, formatCol + 2), xlWs.Cells(lastRow, formatCol + 2)).NumberFormat = "0"
            End If
            
            If formatCol + 3 <= lastCol Then
                ' Colonne "%" â†’ format pourcentage entier (mais les valeurs sont dÃ©jÃ  en % textuel)
                ' On garde le format texte car les valeurs contiennent dÃ©jÃ  le symbole %
                ' Si on voulait un vrai format pourcentage : xlWs.Range(...).NumberFormat = "0%"
            End If
            
            colorIndex = colorIndex + 1
        Next formatCol
    End If
End Sub

' ===================================================================
' FONCTIONS POUR BOUTON ET VISIBILITE DANS "PERSONNALISER LE RUBAN"
' ===================================================================

' Bouton "Ruban" (via Personnaliser le ruban > Macros)
Public Sub ExportMeca_Bouton()
    ' Lance l'export complet (2 onglets, logique existante)
    ExportMecaniqueComplet
End Sub

' CrÃ©e un bouton dans l'onglet ComplÃ©ments (Add-Ins) pour lancer l'export
Public Sub InstallerBoutonExportMeca()
    On Error Resume Next
    ' Nettoyage si dÃ©jÃ  prÃ©sent
    Application.CommandBars("ExportMeca").Delete
    On Error GoTo 0

    Dim cb As CommandBar
    Dim btn As CommandBarButton

    Set cb = Application.CommandBars.Add(Name:="ExportMeca", Position:=msoBarTop, Temporary:=True)
    Set btn = cb.Controls.Add(Type:=msoControlButton)

    With btn
        .Caption = "Export Mecanique"
        .OnAction = "ExportMeca_Bouton"  ' appelle le wrapper
        .Style = msoButtonIconAndCaption
        .FaceId = 176  ' icone standard Office ; changeable si besoin
        .TooltipText = "Exporter Recapitulatif + Donnees detaillees (Mecanique)"
    End With

    cb.Visible = True
End Sub

' Optionnel : suppression du bouton Complements
Public Sub SupprimerBoutonExportMeca()
    On Error Resume Next
    Application.CommandBars("ExportMeca").Delete
    On Error GoTo 0
End Sub

' === FONCTIONS D'INSERTION DES LOGOS ===
Sub InsererLogos(ws As Object)
    Debug.Print "Start: Insertion logos | " & Format(Now, "hh:nn:ss")
    
    ' Insérer le logo Omexom à gauche
    Call InsererLogoOmexom(ws)
    
    ' Insérer le logo du client à droite
    Call InsererLogoClient(ws)
    
    Debug.Print "Done: Insertion logos | " & Format(Now, "hh:nn:ss")
End Sub

Sub InsererLogoOmexom(ws As Object)
    Dim base64Image As String
    Dim byteData() As Byte
    Dim xml As Object, node As Object, stream As Object
    Dim tempFile As String

    On Error GoTo ErrorLogo
    
    Debug.Print "Start: Insertion logo Omexom | " & Format(Now, "hh:nn:ss")
    
    ' === Logo Omexom en base64 ===
    base64Image = GetBase64()
    Debug.Print "Logo Omexom: Base64 recuperee"

    ' Conversion Base64 ? octets
    Set xml = CreateObject("MSXML2.DOMDocument.6.0")
    Set node = xml.createElement("b64")
    node.DataType = "bin.base64"
    node.Text = base64Image
    byteData = node.nodeTypedValue
    Debug.Print "Logo Omexom: Conversion Base64 realisee"

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
    Debug.Print "Logo Omexom: Fichier temporaire cree: " & tempFile

    ' Insertion du logo Omexom (à gauche)
    ws.Shapes.AddPicture tempFile, _
        LinkToFile:=False, _
        SaveWithDocument:=True, _
        Left:=10, Top:=5, Width:=120, Height:=40
    Debug.Print "Logo Omexom: Image insérée dans Excel"

    ' Suppression du fichier temporaire
    On Error Resume Next: Kill tempFile
    Debug.Print "Done: Insertion logo Omexom | " & Format(Now, "hh:nn:ss")
    Exit Sub
    
ErrorLogo:
    Debug.Print "ERREUR insertion logo Omexom | " & Format(Now, "hh:nn:ss") & " | Err=" & Err.Number & " - " & Err.Description
End Sub

Sub InsererLogoClient(ws As Object)
    Dim base64Image As String
    Dim byteData() As Byte
    Dim xml As Object, node As Object, stream As Object
    Dim tempFile As String

    On Error GoTo ErrorLogo
    
    Debug.Print "Start: Insertion logo Client | " & Format(Now, "hh:nn:ss")
    
    ' === Logo Client en base64 ===
    base64Image = GetBase64Client()
    Debug.Print "Logo Client: Base64 récupéré"

    ' Conversion Base64 ? octets
    Set xml = CreateObject("MSXML2.DOMDocument.6.0")
    Set node = xml.createElement("b64")
    node.DataType = "bin.base64"
    node.Text = base64Image
    byteData = node.nodeTypedValue
    Debug.Print "Logo Client: Conversion Base64 réalisée"

    ' Fichier temporaire
    tempFile = Environ$("TEMP") & "\client_logo.png"
    Set stream = CreateObject("ADODB.Stream")
    With stream
        .Type = 1
        .Open
        .Write byteData
        .SaveToFile tempFile, 2
        .Close
    End With
    Debug.Print "Logo Client: Fichier temporaire créé: " & tempFile

    ' Insertion du logo Client (à droite)
    ws.Shapes.AddPicture tempFile, _
        LinkToFile:=False, _
        SaveWithDocument:=True, _
        Left:=140, Top:=5, Width:=120, Height:=40
    Debug.Print "Logo Client: Image insérée dans Excel"

    ' Suppression du fichier temporaire
    On Error Resume Next: Kill tempFile
    Debug.Print "Done: Insertion logo Client | " & Format(Now, "hh:nn:ss")
    Exit Sub
    
ErrorLogo:
    Debug.Print "ERREUR insertion logo Client | " & Format(Now, "hh:nn:ss") & " | Err=" & Err.Number & " - " & Err.Description
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
Function GetBase64Client() As String
    Dim parts(7) As String

    parts(0) = "iVBORw0KGgoAAAANSUhEUgAAAcIAAACWBAMAAAC/XpUjAAAAMFBMVEVHcEzrZgjrZgjrZgjrZgj////2s4XxjUb72L/uchn/9O/0oWfvgzYAAAAAAAAAAABEd2Z9AAAAEHRSTlMAgUC/////////////AAAA6H3hgwAAAAlwSFlzAAALEgAACxIB0t1+/AAAA+RJREFUe"
    parts(1) = "JztnN1S00AUxxfHB1DHB1DHB/DCB/BjoZT2ksUSvGwr6i3QwfEyYHW8rMIDEGRGL+n4AnV0eCqzm03SfCxN11RmD//fVZPuSfLj5OOcbgtjAAAAAAAAAAAAAAD8J25bcMc68h5j9y3CbjG2ZBEWHSe34LkMXLKJfMLYY4uwp4w9tNldZDiYm33HDMXcbMDQDAwNwNAEDGFoAQwNwNAEDI"
    parts(2) = "t0AsKG3rD/c8JbAU3D4eGrfR3TI2d4dvj2/XTMLiXDzrA/GOVjqBjKi26nNGZMwdCTF10heRoS16FfPjpyPqBgWDw924P+MDhXL+s0bMz4tCQ9j2o2nEyPaQ7629JKfI+WxUzDXTWkiuGy8SaX/1NnDFfUm/9imGz49972p+TijLzbJAw/hm819vrDzMofUUiDhKEIL7r8KpXCI7119w1"
    parts(3) = "LkClsfyBs2JG3tYtNzlepGspLsx2sc75G1LAzCUdfyFJgmaihursG0nCXpmFHDh4L0SVrGJrxVrTHXu2Gnb4RQ9VWv2GcQml4ULuhaWSGBRt2eZRCuVGShlsjbebxqPAmZyi7pq9BdLY2KRpuyZEH+lgaFA1VCoU+lhZBw2/huKZKoXgRb5yUoeqa1qLXL3XhTctQdk06hWKdoqFK4bFe"
    parts(4) = "8Dn/Qs7wM5ddU2r4i5ph1PjGS1199JQMMymUz40eMcNsCsPCu0nNUDa+rXRxpGsbOoZq3+N0OVwK6jcUQyOTRRv6PJNC2SaKBRiaWXSPn09hUniTMXydbEx7jeKMEjFMu6Z0xQopw3OeTaFsLVYpGRZSmBbeNAw9mcLjzBTUSXJIJAynuyaNH7cWJAwzXZOmGxfeJAzVjG9umrSbPB0JG"
    parts(5) = "KoU/inur7cAw6upS30+3TVpksKbgKGeLszBk/PWfcNoxjevzePC231DlcJ3hbX6yzQUDH2e6ZpiraS1cN4w3zVp0sLbdUMv99lFTFp4u26YzPjm2CRjKEvuVVFkPZm/cNxQdU29EsOtpCx13NCUQuHtNOJHpNOGxhRO47KhJ8v5tbwRJcOSxpeWYVnju2hDz/I7UYfGsPNLDKul8Ko74N"
    parts(6) = "Ib4RS+2TA310TQMDtdSNCwagpdNayaQmcNVQrHFQSdNZRzTSVdEx3Ds8opdNVQrq+WQkcN1SHNKrmdNjR2TXMbHql/tFTF0PL3h+0ZYZNSQ9k1Nd/UYngpGcPq1PALyxOem/ElZ+hXKrldNjyp0jU5bbgxRwrdNBSnVW8zzhrOAwwvAYYGYGgChjC0AIYGYGgChjC04EoMLbgmhjdGswc"
    parts(7) = "WeMTYXYuwZ4zdtAhTx8ke2MAAAAAAAAAAAAAAAADXjL94K0pTDTqoxgAAAABJRU5ErkJggg=="

    GetBase64Client = Join(parts, "")
End Function









