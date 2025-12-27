Option Explicit

' =============================================================================
' ONE-FILE REPORT GENERATOR (MS Project -> Word .docx)
' Entry point: BuildWeeklyReport
' =============================================================================

' =========================
' CONFIG
' =========================
Private Const REPORT_TITLE As String = "OMEXOM/EDF RE RAPPORT HEBDOMADAIRE"
Private Const PROJECT_NAME As String = "CHANTIER PHOTOVOLTAÏQUE 130MWc PREVENCHERES"

Private Const BASE_PROJECT_PATH As String = "\VINCI Energies\GO-ENR SO - Affaires - EDF - PREVENCHERES - P.0853898.T.01\05 - CHANTIER"
Private Const BASE_SUIVI_LOGISTIQUE As String = BASE_PROJECT_PATH & "\K - SUIVI LOGISTIQUE"

' Chemin master MPP: utilise %USERPROFILE% pour être accessible à tous les utilisateurs
' Le chemin sera résolu dynamiquement via GetMasterProjectPath()
Private Const MASTER_PROJECT_RELATIVE_PATH As String = "\VINCI Energies\GO-ENR SO - Affaires - General\EDF - PREVENCHERES - P.0853898.T.01\05 - CHANTIER\G - PLANNING\Prévenchères Planning Maître\Prevencheres_2026_Maitre.mpp"

' Word Late Binding constants
Private Const WD_COLLAPSE_END As Long = 0
Private Const WD_PAGE_BREAK As Long = 7
Private Const WD_ALIGN_LEFT As Long = 0
Private Const WD_ALIGN_CENTER As Long = 1

' =========================
' PUBLIC ENTRY POINT
' =========================
Public Sub BuildWeeklyReport()
    Dim wdApp As Object
    Dim doc As Object
    Dim outFolder As String
    Dim outPath As String

    On Error GoTo EH

    outFolder = GetDefaultOutputFolder()
    If Not EnsureFolder(outFolder) Then
        MsgBox "Impossible de créer/accéder au dossier de sortie:" & vbCrLf & outFolder, vbCritical
        Exit Sub
    End If

    outPath = outFolder & "\Rapport_Hebdo_Prevencheres_" & GetReportDateTimeStamp() & ".docx"

    Set wdApp = WordAppCreate()
    If wdApp Is Nothing Then Exit Sub

    Set doc = WordDocCreate(wdApp)
    If doc Is Nothing Then
        wdApp.Quit
        Exit Sub
    End If

    ' ===== Sections =====
    Section1_CoverPage doc
    Section2_Avancement doc
    Section3_Qualite doc
    Section4_Mobilisations doc
    Section5_HSE doc
    Section6_MasterProject doc
    Section7_SuiviActions doc
    Section8_Divers doc

    WordSaveAsDocx doc, outPath
    WordClose wdApp, doc, True

    MsgBox "Rapport généré:" & vbCrLf & outPath, vbInformation
    Exit Sub

EH:
    MsgBox "Erreur génération rapport:" & vbCrLf & Err.Number & " - " & Err.Description, vbCritical
    On Error Resume Next
    If Not doc Is Nothing Then doc.Close False
    If Not wdApp Is Nothing Then wdApp.Quit
End Sub

' =========================
' SECTIONS
' =========================
Private Sub Section1_CoverPage(ByVal doc As Object)
    AddHeading doc, "1 : Page de Garde", 1
    AddParagraph doc, REPORT_TITLE, WD_ALIGN_CENTER
    AddParagraph doc, "DATE : " & GetReportDateText(), WD_ALIGN_CENTER
    AddParagraph doc, PROJECT_NAME, WD_ALIGN_CENTER
    AddBlankLine doc
    AddParagraph doc, "Photo du chantier, noms des responsables, etc.", WD_ALIGN_LEFT

    ' Exemple image (à brancher)
    ' AddImage doc, "C:\temp\photo.jpg", 450

    AddPageBreak doc
End Sub

Private Sub Section2_Avancement(ByVal doc As Object)
    AddHeading doc, "2 : Etat d'avancement du projet", 1
    AddParagraph doc, "Map sitemark", WD_ALIGN_LEFT
    AddParagraph doc, "Graphiques d'avancement par Zone et Métier", WD_ALIGN_LEFT
    AddBlankLine doc

    ' Graphique 1 : Avancement par ZONE et MÉTIER (% tâches)
    AddHeading doc, "2.1 : Avancement par Zone (% tâches)", 2
    CreateProgressChartByZoneAndMetier doc, "Zone", True
    AddBlankLine doc

    ' Graphique 2 : Avancement par ZONE et MÉTIER (% ressources)
    AddHeading doc, "2.2 : Avancement par Zone (% ressources)", 2
    CreateProgressChartByZoneAndMetier doc, "Zone", False
    AddBlankLine doc

    ' Graphique 3 : Avancement par SOUS-ZONE et MÉTIER (% tâches)
    AddHeading doc, "2.3 : Avancement par Sous-Zone (% tâches)", 2
    CreateProgressChartByZoneAndMetier doc, "SousZone", True
    AddBlankLine doc

    ' Graphique 4 : Avancement par SOUS-ZONE et MÉTIER (% ressources)
    AddHeading doc, "2.4 : Avancement par Sous-Zone (% ressources)", 2
    CreateProgressChartByZoneAndMetier doc, "SousZone", False

    AddPageBreak doc
End Sub

' =========================
' HELPERS SECTION 2 - AVANCEMENT PAR ZONE ET MÉTIER (4 GRAPHIQUES)
' =========================
Private Sub CreateProgressChartByZoneAndMetier(ByVal doc As Object, ByVal groupBy As String, ByVal useTaskPercent As Boolean)
    ' Crée un graphique d'avancement multi-séries (Zone/Sous-Zone × Métier)
    ' groupBy : "Zone" ou "SousZone"
    ' useTaskPercent : True = % tâches, False = % ressources
    
    Dim data As Object
    Dim zones As Object
    Dim metiers As Object
    Dim chartTitle As String
    
    On Error GoTo EH
    
    ' Extraction des données
    Set data = ExtractProgressData(groupBy, useTaskPercent, zones, metiers)
    
    ' Vérifier qu'on a des données
    If data Is Nothing Or data.Count = 0 Then
        AddParagraph doc, "[Aucune donnée disponible pour ce graphique]", WD_ALIGN_LEFT
        Exit Sub
    End If
    
    ' Titre du graphique
    If groupBy = "Zone" Then
        If useTaskPercent Then
            chartTitle = "Avancement par Zone et Métier (% tâches)"
        Else
            chartTitle = "Avancement par Zone et Métier (% ressources)"
        End If
    Else
        If useTaskPercent Then
            chartTitle = "Avancement par Sous-Zone et Métier (% tâches)"
        Else
            chartTitle = "Avancement par Sous-Zone et Métier (% ressources)"
        End If
    End If
    
    ' Création du graphique
    AddMultiSeriesChart doc, data, zones, metiers, chartTitle
    
    Exit Sub
    
EH:
    AddParagraph doc, "[Erreur création graphique: " & Err.Description & "]", WD_ALIGN_LEFT
End Sub

Private Function ExtractProgressData(ByVal groupBy As String, ByVal useTaskPercent As Boolean, ByRef zonesOut As Object, ByRef metiersOut As Object) As Object
    ' Extrait les données d'avancement groupées par (Zone/SousZone, Métier)
    ' Retourne un Dictionary avec clés "Zone|Métier" -> pourcentage
    ' Remplit zonesOut et metiersOut avec les valeurs uniques (Dictionaries)
    
    Dim data As Object
    Dim workDict As Object
    Dim actualWorkDict As Object
    Dim percentDict As Object
    Dim countDict As Object
    Dim t As Task
    Dim zone As String
    Dim metier As String
    Dim key As String
    Dim workMinutes As Double
    Dim actualWorkMinutes As Double
    Dim pctComplete As Double
    Dim finalPercent As Double
    
    ' Compteurs logs
    Dim totalTasks As Long
    Dim ignoredSummary As Long
    Dim ignoredNoWork As Long
    Dim ignoredNoZone As Long
    Dim ignoredNoMetier As Long
    Dim processedTasks As Long
    
    On Error GoTo EH
    
    Debug.Print "=== DEBUT ExtractProgressData (groupBy=" & groupBy & ", useTaskPercent=" & useTaskPercent & ") ==="
    
    Set data = CreateObject("Scripting.Dictionary")
    Set workDict = CreateObject("Scripting.Dictionary")
    Set actualWorkDict = CreateObject("Scripting.Dictionary")
    Set percentDict = CreateObject("Scripting.Dictionary")
    Set countDict = CreateObject("Scripting.Dictionary")
    Set zonesOut = CreateObject("Scripting.Dictionary")
    Set metiersOut = CreateObject("Scripting.Dictionary")
    
    ' Vérification ActiveProject
    If ActiveProject Is Nothing Then
        Debug.Print "ERREUR: ActiveProject est Nothing"
        Set ExtractProgressData = data
        Exit Function
    End If
    
    Debug.Print "ActiveProject OK - Nombre de tâches: " & ActiveProject.Tasks.Count
    
    ' Parcours des tâches
    For Each t In ActiveProject.Tasks
        totalTasks = totalTasks + 1
        
        If Not t Is Nothing Then
            ' Ignorer les Summary tasks
            If t.Summary Then
                ignoredSummary = ignoredSummary + 1
                GoTo NextTask
            End If
            
            ' Récupérer Work
            On Error Resume Next
            workMinutes = t.Work
            If Err.Number <> 0 Or workMinutes = 0 Then
                On Error GoTo EH
                ignoredNoWork = ignoredNoWork + 1
                GoTo NextTask
            End If
            On Error GoTo EH
            
            ' Récupérer Zone ou SousZone selon groupBy
            On Error Resume Next
            If groupBy = "Zone" Then
                zone = Trim(CStr(t.Text2))
            Else ' SousZone
                zone = Trim(CStr(t.Text3))
            End If
            If Err.Number <> 0 Or Len(zone) = 0 Then
                On Error GoTo EH
                ignoredNoZone = ignoredNoZone + 1
                GoTo NextTask
            End If
            On Error GoTo EH
            
            ' Récupérer Métier (Text4)
            On Error Resume Next
            metier = Trim(CStr(t.Text4))
            If Err.Number <> 0 Or Len(metier) = 0 Then
                On Error GoTo EH
                ignoredNoMetier = ignoredNoMetier + 1
                GoTo NextTask
            End If
            On Error GoTo EH
            
            ' Récupérer ActualWork et PercentComplete
            On Error Resume Next
            actualWorkMinutes = t.ActualWork
            If Err.Number <> 0 Then actualWorkMinutes = 0
            
            pctComplete = t.PercentComplete
            If Err.Number <> 0 Then pctComplete = 0
            On Error GoTo EH
            
            ' Clé : "Zone|Métier"
            key = zone & "|" & metier
            
            ' Accumulation
            If Not workDict.Exists(key) Then
                workDict(key) = 0
                actualWorkDict(key) = 0
                percentDict(key) = 0
                countDict(key) = 0
            End If
            
            workDict(key) = workDict(key) + workMinutes
            actualWorkDict(key) = actualWorkDict(key) + actualWorkMinutes
            percentDict(key) = percentDict(key) + pctComplete
            countDict(key) = countDict(key) + 1
            
            ' Enregistrer zone et métier uniques
            If Not zonesOut.Exists(zone) Then zonesOut(zone) = True
            If Not metiersOut.Exists(metier) Then metiersOut(metier) = True
            
            processedTasks = processedTasks + 1
            Debug.Print "  Tâche [" & t.Name & "] - Zone=" & zone & " | Métier=" & metier & " | Work=" & workMinutes & " | ActualWork=" & actualWorkMinutes & " | Pct=" & pctComplete
        End If
        
NextTask:
    Next t
    
    ' Logs récapitulatifs
    Debug.Print "=== RECAPITULATIF ==="
    Debug.Print "Total tâches parcourues: " & totalTasks
    Debug.Print "  - Ignorées (Summary): " & ignoredSummary
    Debug.Print "  - Ignorées (pas de Work): " & ignoredNoWork
    Debug.Print "  - Ignorées (pas de Zone): " & ignoredNoZone
    Debug.Print "  - Ignorées (pas de Métier): " & ignoredNoMetier
    Debug.Print "  - TRAITEES avec succès: " & processedTasks
    Debug.Print "Zones uniques: " & zonesOut.Count
    Debug.Print "Métiers uniques: " & metiersOut.Count
    
    ' Calcul final par clé
    Debug.Print "=== CALCUL FINAL PAR (ZONE|METIER) ==="
    Dim k As Variant
    For Each k In workDict.Keys
        If useTaskPercent Then
            ' Moyenne des PercentComplete
            If countDict(k) > 0 Then
                finalPercent = percentDict(k) / countDict(k)
            Else
                finalPercent = 0
            End If
        Else
            ' (ActualWork / Work) * 100
            If workDict(k) > 0 Then
                finalPercent = (actualWorkDict(k) / workDict(k)) * 100
            Else
                finalPercent = 0
            End If
        End If
        
        data(k) = finalPercent
        Debug.Print k & " => " & Format(finalPercent, "0.00") & "%"
    Next k
    
    Debug.Print "=== FIN ExtractProgressData ===" & vbCrLf
    
    Set ExtractProgressData = data
    Exit Function
    
EH:
    Debug.Print "ERREUR dans ExtractProgressData: " & Err.Number & " - " & Err.Description
    Set ExtractProgressData = data
End Function

Private Sub AddMultiSeriesChart(ByVal doc As Object, ByVal data As Object, ByVal zones As Object, ByVal metiers As Object, ByVal chartTitle As String)
    ' Crée un graphique en colonnes groupées (multi-séries) dans Word
    ' Structure Excel : lignes = zones, colonnes = métiers
    
    Dim chartShape As Object
    Dim chart As Object
    Dim chartData As Object
    Dim workbook As Object
    Dim worksheet As Object
    Dim rng As Object
    Dim zonesArray() As String
    Dim metiersArray() As String
    Dim i As Long, j As Long
    Dim zone As Variant
    Dim metier As Variant
    Dim key As String
    Dim value As Double
    Dim rowIdx As Long
    Dim colIdx As Long
    Dim errNum As Long
    Dim errDesc As String
    
    On Error GoTo EH
    
    Debug.Print "=== DEBUT AddMultiSeriesChart ==="
    
    ' Conversion des dictionnaires en tableaux
    Debug.Print "Conversion zones..."
    ReDim zonesArray(zones.Count - 1)
    i = 0
    For Each zone In zones.Keys
        zonesArray(i) = CStr(zone)
        Debug.Print "  Zone[" & i & "] = " & zonesArray(i)
        i = i + 1
    Next zone
    
    Debug.Print "Conversion métiers..."
    ReDim metiersArray(metiers.Count - 1)
    i = 0
    For Each metier In metiers.Keys
        metiersArray(i) = CStr(metier)
        Debug.Print "  Métier[" & i & "] = " & metiersArray(i)
        i = i + 1
    Next metier
    
    Debug.Print "=== CREATION GRAPHIQUE ==="
    Debug.Print "Zones: " & (UBound(zonesArray) + 1)
    Debug.Print "Métiers: " & (UBound(metiersArray) + 1)
    
    ' Positionner le curseur à la fin du document
    Debug.Print "Positionnement curseur..."
    Set rng = doc.Range(doc.Content.End - 1)
    rng.InsertAfter vbCrLf
    rng.Collapse WD_COLLAPSE_END
    Debug.Print "Curseur positionné OK"
    
    ' Créer le graphique (type 51 = colonnes groupées)
    ' Note : AddChart nécessite Excel installé, sinon on crée un tableau
    Debug.Print "Création du graphique Word (type 51)..."
    
    On Error Resume Next
    Set chartShape = rng.InlineShapes.AddChart(51)
    errNum = Err.Number
    errDesc = Err.Description
    Err.Clear
    On Error GoTo EH
    
    If errNum <> 0 Then
        Debug.Print "AVERTISSEMENT: Impossible de créer un graphique (Erreur " & errNum & ": " & errDesc & ")"
        Debug.Print "Création d'un tableau Word à la place..."
        
        ' Fallback: créer un tableau Word avec les données
        CreateDataTable doc, data, zonesArray, metiersArray, chartTitle
        
        Debug.Print "=== TABLEAU CREE AVEC SUCCES (fallback graphique) ===" & vbCrLf
        Exit Sub
    End If
    
    Debug.Print "chartShape créé OK"
    
    Set chart = chartShape.Chart
    Debug.Print "chart récupéré OK"
    
    ' Titre
    Debug.Print "Ajout du titre..."
    On Error Resume Next
    chart.HasTitle = True
    errNum = Err.Number
    errDesc = Err.Description
    Err.Clear
    On Error GoTo EH
    If errNum <> 0 Then
        Debug.Print "AVERTISSEMENT: Impossible de définir HasTitle: " & errNum & " - " & errDesc
    End If
    
    On Error Resume Next
    chart.ChartTitle.Text = chartTitle
    errNum = Err.Number
    errDesc = Err.Description
    Err.Clear
    On Error GoTo EH
    If errNum <> 0 Then
        Debug.Print "AVERTISSEMENT: Impossible de définir ChartTitle: " & errNum & " - " & errDesc
    Else
        Debug.Print "Titre défini OK: " & chartTitle
    End If
    
    ' Accès au workbook de données
    Debug.Print "Accès au ChartData..."
    Set chartData = chart.ChartData
    Debug.Print "chartData récupéré OK"
    
    Debug.Print "Activation du ChartData..."
    chartData.Activate
    Debug.Print "chartData.Activate OK"
    
    Debug.Print "Accès au Workbook..."
    Set workbook = chartData.Workbook
    Debug.Print "workbook récupéré OK"
    
    Debug.Print "Accès à la feuille 1..."
    Set worksheet = workbook.Worksheets(1)
    Debug.Print "worksheet récupéré OK"
    
    ' Effacer données par défaut
    Debug.Print "Effacement des données par défaut..."
    On Error Resume Next
    worksheet.Cells.Clear
    errNum = Err.Number
    errDesc = Err.Description
    Err.Clear
    On Error GoTo EH
    If errNum <> 0 Then
        Debug.Print "AVERTISSEMENT: Impossible d'effacer les cellules: " & errNum & " - " & errDesc
    Else
        Debug.Print "Cellules effacées OK"
    End If
    
    ' Construction du tableau Excel
    Debug.Print "Construction du tableau Excel..."
    ' Ligne 1 : En-têtes (Zone | Métier1 | Métier2 | ...)
    worksheet.Cells(1, 1).Value = "Zone"
    Debug.Print "  En-tête [1,1] = Zone"
    
    For j = 0 To UBound(metiersArray)
        worksheet.Cells(1, j + 2).Value = metiersArray(j)
        Debug.Print "  En-tête [1," & (j + 2) & "] = " & metiersArray(j)
    Next j
    
    ' Lignes suivantes : Une ligne par zone
    Debug.Print "Remplissage des données..."
    For i = 0 To UBound(zonesArray)
        rowIdx = i + 2
        zone = zonesArray(i)
        worksheet.Cells(rowIdx, 1).Value = zone
        Debug.Print "  Ligne " & rowIdx & " - Zone: " & zone
        
        ' Une colonne par métier
        For j = 0 To UBound(metiersArray)
            colIdx = j + 2
            metier = metiersArray(j)
            key = zone & "|" & metier
            
            ' Récupérer la valeur ou 0 si pas de données
            If data.Exists(key) Then
                value = data(key)
            Else
                value = 0
            End If
            
            worksheet.Cells(rowIdx, colIdx).Value = value
            Debug.Print "    [" & rowIdx & "," & colIdx & "] " & key & " = " & Format(value, "0.00") & "%"
        Next j
    Next i
    
    ' Définir la plage de données pour le graphique
    Dim lastRow As Long, lastCol As Long
    lastRow = UBound(zonesArray) + 2
    lastCol = UBound(metiersArray) + 2
    
    Debug.Print "Définition de la plage de données: [1,1] à [" & lastRow & "," & lastCol & "]"
    
    On Error Resume Next
    chart.SetSourceData worksheet.Range(worksheet.Cells(1, 1), worksheet.Cells(lastRow, lastCol))
    errNum = Err.Number
    errDesc = Err.Description
    Err.Clear
    On Error GoTo EH
    
    If errNum <> 0 Then
        Debug.Print "ERREUR SetSourceData: " & errNum & " - " & errDesc
        ' On continue quand même pour fermer le workbook
    Else
        Debug.Print "SetSourceData OK"
    End If
    
    ' Fermer le workbook
    Debug.Print "Fermeture du workbook..."
    workbook.Close False
    Debug.Print "workbook fermé OK"
    
    ' Saut de ligne après graphique
    rng.InsertAfter vbCrLf
    rng.Collapse WD_COLLAPSE_END
    
    Debug.Print "=== GRAPHIQUE CREE AVEC SUCCES ===" & vbCrLf
    
    Exit Sub
    
EH:
    errNum = Err.Number
    errDesc = Err.Description
    Debug.Print "ERREUR AddMultiSeriesChart: " & errNum & " - " & errDesc
    On Error Resume Next
    If Not rng Is Nothing Then
        rng.InsertAfter "[Erreur création graphique: #" & errNum & " - " & errDesc & "]" & vbCrLf
    Else
        AddParagraph doc, "[Erreur création graphique: #" & errNum & " - " & errDesc & "]", WD_ALIGN_LEFT
    End If
    On Error GoTo 0
End Sub

Private Sub CreateDataTable(ByVal doc As Object, ByVal data As Object, ByRef zonesArray() As String, ByRef metiersArray() As String, ByVal tableTitle As String)
    ' Crée un tableau Word avec les données (fallback si graphiques non disponibles)
    ' Structure : lignes = zones, colonnes = métiers
    
    Dim tbl As Object
    Dim rng As Object
    Dim i As Long, j As Long
    Dim zone As String
    Dim metier As String
    Dim key As String
    Dim value As Double
    Dim numRows As Long
    Dim numCols As Long
    
    On Error GoTo EH
    
    Debug.Print "=== CREATION TABLEAU WORD ==="
    
    ' Calculer dimensions du tableau
    numRows = UBound(zonesArray) + 2  ' +1 pour en-tête, +1 pour 0-based
    numCols = UBound(metiersArray) + 2  ' +1 pour colonne Zone, +1 pour 0-based
    
    Debug.Print "Tableau: " & numRows & " lignes × " & numCols & " colonnes"
    
    ' Positionner le curseur
    Set rng = doc.Range(doc.Content.End - 1)
    
    ' Ajouter le titre
    rng.InsertAfter tableTitle & vbCrLf
    rng.Collapse WD_COLLAPSE_END
    
    ' Créer le tableau
    Set tbl = doc.Tables.Add(rng, numRows, numCols)
    
    ' Style du tableau
    On Error Resume Next
    tbl.Style = "Grille du tableau moyenne 2"  ' Style par défaut
    tbl.AutoFitBehavior 2  ' wdAutoFitContent = 2
    On Error GoTo EH
    
    ' En-tête (ligne 1)
    tbl.Cell(1, 1).Range.Text = "Zone \ Métier"
    tbl.Cell(1, 1).Range.Bold = True
    tbl.Cell(1, 1).Shading.BackgroundPatternColor = RGB(200, 200, 200)
    
    For j = 0 To UBound(metiersArray)
        tbl.Cell(1, j + 2).Range.Text = metiersArray(j)
        tbl.Cell(1, j + 2).Range.Bold = True
        tbl.Cell(1, j + 2).Shading.BackgroundPatternColor = RGB(200, 200, 200)
        Debug.Print "  En-tête col " & (j + 2) & ": " & metiersArray(j)
    Next j
    
    ' Données (lignes suivantes)
    For i = 0 To UBound(zonesArray)
        zone = zonesArray(i)
        
        ' Colonne 1 : nom de la zone
        tbl.Cell(i + 2, 1).Range.Text = zone
        tbl.Cell(i + 2, 1).Range.Bold = True
        
        ' Colonnes suivantes : valeurs par métier
        For j = 0 To UBound(metiersArray)
            metier = metiersArray(j)
            key = zone & "|" & metier
            
            ' Récupérer la valeur
            If data.Exists(key) Then
                value = data(key)
            Else
                value = 0
            End If
            
            ' Formater avec 1 décimale et symbole %
            tbl.Cell(i + 2, j + 2).Range.Text = Format(value, "0.0") & "%"
            
            ' Colorier selon la valeur (vert pour >50%, jaune pour 25-50%, rouge pour <25%)
            On Error Resume Next
            If value >= 50 Then
                tbl.Cell(i + 2, j + 2).Shading.BackgroundPatternColor = RGB(200, 255, 200)  ' Vert clair
            ElseIf value >= 25 Then
                tbl.Cell(i + 2, j + 2).Shading.BackgroundPatternColor = RGB(255, 255, 200)  ' Jaune clair
            ElseIf value > 0 Then
                tbl.Cell(i + 2, j + 2).Shading.BackgroundPatternColor = RGB(255, 200, 200)  ' Rouge clair
            End If
            On Error GoTo EH
            
            Debug.Print "  [" & (i + 2) & "," & (j + 2) & "] " & zone & "|" & metier & " = " & Format(value, "0.0") & "%"
        Next j
    Next i
    
    ' Centrer les valeurs
    On Error Resume Next
    tbl.Range.ParagraphFormat.alignment = WD_ALIGN_CENTER
    On Error GoTo EH
    
    ' Saut de ligne après le tableau
    rng.InsertAfter vbCrLf
    rng.Collapse WD_COLLAPSE_END
    
    Debug.Print "=== TABLEAU CREE AVEC SUCCES ===" & vbCrLf
    
    Exit Sub
    
EH:
    Debug.Print "ERREUR CreateDataTable: " & Err.Number & " - " & Err.Description
    On Error Resume Next
    AddParagraph doc, "[Erreur création tableau: " & Err.Description & "]", WD_ALIGN_LEFT
    On Error GoTo 0
End Sub

Private Sub Section3_Qualite(ByVal doc As Object)
    Dim col As Collection, i As Long

    AddHeading doc, "3 : Suivi des contrôles qualité", 1
    AddParagraph doc, "Extract depuis Teepee ou MS ? ==> Taches dans MS", WD_ALIGN_LEFT
    AddParagraph doc, "BI Teepee: pas pour le moment", WD_ALIGN_LEFT
    AddBlankLine doc

    Set col = GetQualityTasks()

    If col.Count = 0 Then
        AddParagraph doc, "Aucune tâche qualité détectée (filtre actuel: 'qualit' ou 'QC').", WD_ALIGN_LEFT
    Else
        AddParagraph doc, "Tâches qualité détectées: " & col.Count, WD_ALIGN_LEFT
        For i = 1 To col.Count
            AddParagraph doc, "• " & col(i).Name, WD_ALIGN_LEFT
        Next i
    End If

    AddPageBreak doc
End Sub

Private Sub Section4_Mobilisations(ByVal doc As Object)
    AddHeading doc, "4 : Mobilisations sur site", 1
    AddParagraph doc, "Arrivées/départs de sous-traitants:", WD_ALIGN_LEFT
    AddParagraph doc, "[À compléter]", WD_ALIGN_LEFT
    AddBlankLine doc
    AddParagraph doc, "Nombre de personnels sur site avec prévisionnel (courbe):", WD_ALIGN_LEFT
    AddParagraph doc, "[À brancher] Source Excel/Project", WD_ALIGN_LEFT
    AddPageBreak doc
End Sub

Private Sub Section5_HSE(ByVal doc As Object)
    AddHeading doc, "5 : HSE", 1
    AddParagraph doc, "Zone de texte à incrémenter (incidents/observations)", WD_ALIGN_LEFT
    AddParagraph doc, "[À brancher] Source CR réunions / outil HSE", WD_ALIGN_LEFT
    AddPageBreak doc
End Sub

Private Sub Section6_MasterProject(ByVal doc As Object)
    AddHeading doc, "6 : MS Project maître", 1
    AddParagraph doc, "Total + zoom à 3 semaines", WD_ALIGN_LEFT
    AddBlankLine doc
    AddParagraph doc, "Fichier maître:", WD_ALIGN_LEFT
    AddParagraph doc, GetMasterProjectPath(), WD_ALIGN_LEFT
    AddBlankLine doc
    AddParagraph doc, "[À brancher] Export image/PDF du Gantt et insertion", WD_ALIGN_LEFT
    AddPageBreak doc
End Sub

Private Sub Section7_SuiviActions(ByVal doc As Object)
    AddHeading doc, "7 : Suivi d'actions", 1
    AddParagraph doc, "Extract Excel suivi d'actions OMX - EDF", WD_ALIGN_LEFT
    AddParagraph doc, "[À brancher] Lecture Excel + insertion tableau Word", WD_ALIGN_LEFT
    AddPageBreak doc
End Sub

Private Sub Section8_Divers(ByVal doc As Object)
    AddHeading doc, "8 : Divers", 1
    AddParagraph doc, "Intégrations de photos, événements sur site (visites…)", WD_ALIGN_LEFT
    AddParagraph doc, "[À compléter]", WD_ALIGN_LEFT
End Sub

' =========================
' PROJECT DATA (MS Project)
' =========================
Private Function GetQualityTasks() As Collection
    Dim col As New Collection
    Dim t As Task

    On Error GoTo EH

    If ActiveProject Is Nothing Then
        Set GetQualityTasks = col
        Exit Function
    End If

    For Each t In ActiveProject.Tasks
        If Not t Is Nothing Then
            If InStr(1, t.Name, "qualit", vbTextCompare) > 0 Or InStr(1, t.Name, "QC", vbTextCompare) > 0 Then
                col.Add t
            End If
        End If
    Next t

    Set GetQualityTasks = col
    Exit Function

EH:
    Set GetQualityTasks = col
End Function

' =========================
' WORD HELPERS (Late Binding)
' =========================
Private Function WordAppCreate() As Object
    Dim wdApp As Object
    On Error GoTo EH

    Set wdApp = CreateObject("Word.Application")
    wdApp.Visible = False
    Set WordAppCreate = wdApp
    Exit Function

EH:
    MsgBox "Impossible de créer Word: " & Err.Description, vbCritical
    Set WordAppCreate = Nothing
End Function

Private Function WordDocCreate(ByVal wdApp As Object) As Object
    On Error GoTo EH
    Set WordDocCreate = wdApp.Documents.Add
    Exit Function
EH:
    MsgBox "Impossible de créer le document Word: " & Err.Description, vbCritical
    Set WordDocCreate = Nothing
End Function

Private Sub WordSaveAsDocx(ByVal doc As Object, ByVal fullPath As String)
    On Error GoTo EH
    doc.SaveAs2 fullPath
    Exit Sub
EH:
    MsgBox "Erreur sauvegarde docx: " & Err.Description, vbCritical
End Sub

Private Sub WordClose(ByVal wdApp As Object, ByVal doc As Object, Optional ByVal saveChanges As Boolean = True)
    On Error Resume Next
    If Not doc Is Nothing Then doc.Close saveChanges
    If Not wdApp Is Nothing Then wdApp.Quit
    On Error GoTo 0
End Sub

Private Sub AddHeading(ByVal doc As Object, ByVal text As String, Optional ByVal level As Long = 1)
    Dim rng As Object
    On Error GoTo EH

    Set rng = doc.Range(doc.Content.End - 1)
    rng.InsertAfter text & vbCrLf
    rng.Style = "Heading " & CStr(level)
    rng.Collapse WD_COLLAPSE_END
    Exit Sub

EH:
    ' Si le style Heading n'existe pas, on retombe en Normal sans planter
    On Error Resume Next
    If Not rng Is Nothing Then rng.Style = "Normal"
    On Error GoTo 0
End Sub

Private Sub AddParagraph(ByVal doc As Object, ByVal text As String, Optional ByVal alignment As Long = WD_ALIGN_LEFT)
    Dim rng As Object
    Set rng = doc.Range(doc.Content.End - 1)
    rng.InsertAfter text & vbCrLf
    rng.ParagraphFormat.alignment = alignment
    rng.Collapse WD_COLLAPSE_END
End Sub

Private Sub AddBlankLine(ByVal doc As Object)
    AddParagraph doc, "", WD_ALIGN_LEFT
End Sub

Private Sub AddPageBreak(ByVal doc As Object)
    On Error Resume Next
    doc.Range(doc.Content.End - 1).InsertBreak WD_PAGE_BREAK
    On Error GoTo 0
End Sub

Private Sub AddImage(ByVal doc As Object, ByVal imagePath As String, Optional ByVal widthPoints As Double = 420)
    Dim rng As Object

    If Len(Dir(imagePath)) = 0 Then
        AddParagraph doc, "[Image introuvable] " & imagePath, WD_ALIGN_LEFT
        Exit Sub
    End If

    Set rng = doc.Range(doc.Content.End - 1)
    rng.InlineShapes.AddPicture imagePath
    On Error Resume Next
    rng.InlineShapes(rng.InlineShapes.Count).Width = widthPoints
    On Error GoTo 0
    rng.InsertAfter vbCrLf
End Sub

' =========================
' PATHS + UTILS
' =========================
Private Function ZonesList() As Variant
    ZonesList = Array("1", "2", "3A", "3B", "3C", "4", "5")
End Function

Private Function GetMasterProjectPath() As String
    ' Construit le chemin complet du fichier master MPP en utilisant %USERPROFILE%
    ' Compatible avec tous les utilisateurs
    GetMasterProjectPath = ExpandEnv("%USERPROFILE%") & MASTER_PROJECT_RELATIVE_PATH
End Function

Private Function GetDefaultOutputFolder() As String
    GetDefaultOutputFolder = ExpandEnv("%USERPROFILE%") & "\Desktop"
End Function

Private Function ExpandEnv(ByVal envString As String) As String
    ExpandEnv = CreateObject("WScript.Shell").ExpandEnvironmentStrings(envString)
End Function

Private Function EnsureFolder(ByVal folderPath As String) As Boolean
    On Error GoTo EH
    If Len(Dir(folderPath, vbDirectory)) = 0 Then MkDir folderPath
    EnsureFolder = True
    Exit Function
EH:
    EnsureFolder = False
End Function

Private Function GetReportDateText() As String
    GetReportDateText = Format(Date, "dd/mm/yyyy")
End Function

Private Function GetReportDateTimeStamp() As String
    GetReportDateTimeStamp = Format(Now, "yyyy-mm-dd_hhmm")
End Function

' (Optionnel) chemins si tu veux les réutiliser plus tard
Private Function GetPath_FormatOutils() As String
    GetPath_FormatOutils = ExpandEnv("%USERPROFILE%") & BASE_SUIVI_LOGISTIQUE & "\00 - FORMAT DES OUTILS"
End Function

Private Function GetPath_Beton() As String
    GetPath_Beton = ExpandEnv("%USERPROFILE%") & BASE_SUIVI_LOGISTIQUE & "\02 - BETON"
End Function

Private Function GetPath_Structures() As String
    GetPath_Structures = ExpandEnv("%USERPROFILE%") & BASE_SUIVI_LOGISTIQUE & "\01 - STRUCTURES"
End Function

Private Function GetPath_CRReunionTemplate() As String
    GetPath_CRReunionTemplate = ExpandEnv("%USERPROFILE%") & BASE_PROJECT_PATH & "\B - CR REUNION\00 - Template"
End Function

