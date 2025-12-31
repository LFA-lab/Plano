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
Private Const WD_ALIGN_CENTER As Long = 1kl56

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

    ' Générer fichier de traçabilité des données (pour validation)
    ExportProjectDataTrace outFolder

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
    Dim percentDict As Object
    Dim percentWorkDict As Object
    Dim countDict As Object
    Dim t As Task
    Dim zone As String
    Dim metier As String
    Dim key As String
    Dim pctComplete As Double
    Dim pctWorkComplete As Double
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
    Set percentDict = CreateObject("Scripting.Dictionary")
    Set percentWorkDict = CreateObject("Scripting.Dictionary")
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
            
            ' Récupérer Work (juste pour filtrage)
            On Error Resume Next
            If t.Work = 0 Then
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
            
            ' Récupérer PercentComplete et PercentWorkComplete
            On Error Resume Next
            pctComplete = t.PercentComplete
            If Err.Number <> 0 Then pctComplete = 0
            
            pctWorkComplete = t.PercentWorkComplete
            If Err.Number <> 0 Then pctWorkComplete = 0
            On Error GoTo EH
            
            ' Clé : "Zone|Métier"
            key = zone & "|" & metier
            
            ' Accumulation
            If Not percentDict.Exists(key) Then
                percentDict(key) = 0
                percentWorkDict(key) = 0
                countDict(key) = 0
            End If
            
            percentDict(key) = percentDict(key) + pctComplete
            percentWorkDict(key) = percentWorkDict(key) + pctWorkComplete
            countDict(key) = countDict(key) + 1
            
            ' Enregistrer zone et métier uniques
            If Not zonesOut.Exists(zone) Then zonesOut(zone) = True
            If Not metiersOut.Exists(metier) Then metiersOut(metier) = True
            
            processedTasks = processedTasks + 1
            Debug.Print "  Tâche [" & t.Name & "] - Zone=" & zone & " | Métier=" & metier & " | PctComplete=" & pctComplete & " | PctWorkComplete=" & pctWorkComplete
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
    For Each k In percentDict.Keys
        If useTaskPercent Then
            ' Moyenne des PercentComplete
            If countDict(k) > 0 Then
                finalPercent = percentDict(k) / countDict(k)
            Else
                finalPercent = 0
            End If
        Else
            ' Moyenne des PercentWorkComplete (travail et consommables)
            If countDict(k) > 0 Then
                finalPercent = percentWorkDict(k) / countDict(k)
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
    ' Section 3 : Suivi des contrôles qualité
    ' Sépare les CQ sur tâches normales et les CQ dédiées
    
    On Error GoTo EH
    
    AddHeading doc, "3 : Suivi des contrôles qualité", 1
    AddBlankLine doc
    
    ' 3.1 : Tableau CQ sur tâches normales (Text4 <> "CQ")
    AddHeading doc, "3.1 : CQ sur tâches normales - Tableau récapitulatif", 2
    CreateQualityTable doc, False  ' False = tâches normales
    AddBlankLine doc
    
    ' 3.2 : Graphique CQ sur tâches normales
    AddHeading doc, "3.2 : CQ sur tâches normales - Graphique d'avancement", 2
    CreateQualityChart doc, False  ' False = tâches normales
    AddBlankLine doc
    
    ' 3.3 : Tableau CQ dédiées (Text4 = "CQ")
    AddHeading doc, "3.3 : CQ dédiées - Tableau récapitulatif", 2
    CreateQualityTable doc, True  ' True = tâches CQ dédiées
    AddBlankLine doc
    
    ' 3.4 : Graphique CQ dédiées
    AddHeading doc, "3.4 : CQ dédiées - Graphique d'avancement", 2
    CreateQualityChart doc, True  ' True = tâches CQ dédiées
    
    AddPageBreak doc
    Exit Sub
    
EH:
    AddParagraph doc, "[Erreur Section 3: " & Err.Number & " - " & Err.Description & "]", WD_ALIGN_LEFT
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
    
    Dim mppPath As String
    Dim imgPath As String
    
    mppPath = GetMasterProjectPath()
    
    ' Chemin vers l'image pré-exportée (via macro dédiée du fichier maître)
    imgPath = ExpandEnv("%USERPROFILE%") & BASE_PROJECT_PATH & "\G - PLANNING\Gantt_Export_3Semaines.png"
    
    AddParagraph doc, "Fichier maître:", WD_ALIGN_LEFT
    AddParagraph doc, mppPath, WD_ALIGN_LEFT
    AddBlankLine doc
    
    AddParagraph doc, "Vue Gantt (3 semaines):", WD_ALIGN_LEFT
    AddBlankLine doc
    
    ' Insertion de l'image pré-exportée
    If Len(Dir(imgPath)) > 0 Then
        AddImage doc, imgPath, 650
        ' Afficher la date de dernière mise à jour
        On Error Resume Next
        AddParagraph doc, "[Dernière mise à jour: " & Format(FileDateTime(imgPath), "dd/mm/yyyy à hh:nn") & "]", WD_ALIGN_LEFT
        On Error GoTo 0
    Else
        AddParagraph doc, "[Image du Gantt non trouvée]", WD_ALIGN_LEFT
        AddParagraph doc, "Lancez la macro d'export depuis le fichier maître avant de générer le rapport.", WD_ALIGN_LEFT
        AddParagraph doc, "Chemin attendu: " & imgPath, WD_ALIGN_LEFT
    End If
    
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
' HELPERS SECTION 3 - QUALITÉ (CONTRÔLES CQ)
' =========================
Private Function ExtractQualityData(ByVal onlyDedicatedCQ As Boolean, ByRef zonesOut As Object, ByRef metiersOut As Object) As Object
    ' Extrait les données CQ depuis MS Project (tâches avec assignation ressource "CQ")
    ' onlyDedicatedCQ: True = seulement CQ dédiées (Text4="CQ"), False = seulement CQ sur tâches normales (Text4<>"CQ")
    ' Retourne un Dictionary avec clés "Zone|Métier" -> { total, completed, avgPercent }
    ' Remplit zonesOut et metiersOut avec les valeurs uniques (Dictionaries)
    
    Dim data As Object
    Dim totalCountDict As Object
    Dim completedCountDict As Object
    Dim sumPercentDict As Object
    Dim t As Task
    Dim a As Object
    Dim hasCQ As Boolean
    Dim zone As String
    Dim metier As String
    Dim text4Val As String
    Dim key As String
    Dim pct As Double
    Dim cqNormales As Long
    Dim cqDediees As Long
    Dim totalCQ As Long
    Dim k As Variant
    Dim dictItem As Object
    
    On Error GoTo EH
    
    Debug.Print "=== DEBUT ExtractQualityData (onlyDedicatedCQ=" & onlyDedicatedCQ & ") ==="
    
    Set data = CreateObject("Scripting.Dictionary")
    Set totalCountDict = CreateObject("Scripting.Dictionary")
    Set completedCountDict = CreateObject("Scripting.Dictionary")
    Set sumPercentDict = CreateObject("Scripting.Dictionary")
    Set zonesOut = CreateObject("Scripting.Dictionary")
    Set metiersOut = CreateObject("Scripting.Dictionary")
    
    cqNormales = 0
    cqDediees = 0
    
    ' Vérification ActiveProject
    If ActiveProject Is Nothing Then
        Debug.Print "ERREUR: ActiveProject est Nothing"
        Set ExtractQualityData = data
        Exit Function
    End If
    
    Debug.Print "ActiveProject OK - Nombre de tâches: " & ActiveProject.Tasks.Count
    
    ' Parcours des tâches
    For Each t In ActiveProject.Tasks
        If Not t Is Nothing Then
            ' Ignorer les Summary tasks
            If t.Summary Then GoTo NextTaskCQ
            
            ' Vérifier si la tâche a une affectation de ressource "CQ"
            hasCQ = False
            On Error Resume Next
            For Each a In t.Assignments
                If Not a Is Nothing Then
                    If Not a.Resource Is Nothing Then
                        If UCase(Trim(a.Resource.Name)) = "CQ" Then
                            hasCQ = True
                            Exit For
                        End If
                    End If
                End If
            Next a
            On Error GoTo EH
            
            ' Si pas de ressource CQ, ignorer cette tâche
            If Not hasCQ Then GoTo NextTaskCQ
            
            ' Récupérer Zone (Text2)
            On Error Resume Next
            zone = Trim(CStr(t.Text2))
            If Err.Number <> 0 Or Len(zone) = 0 Then
                On Error GoTo EH
                GoTo NextTaskCQ
            End If
            On Error GoTo EH
            
            ' Récupérer Text4 pour déterminer le type de CQ
            On Error Resume Next
            text4Val = Trim(CStr(t.Text4))
            If Err.Number <> 0 Then text4Val = ""
            On Error GoTo EH
            
            ' Filtrer selon le type demandé
            If onlyDedicatedCQ Then
                ' On veut seulement les CQ dédiées (Text4 = "CQ")
                If UCase(text4Val) <> "CQ" Then GoTo NextTaskCQ
            Else
                ' On veut seulement les CQ sur tâches normales (Text4 <> "CQ")
                If UCase(text4Val) = "CQ" Or Len(text4Val) = 0 Then GoTo NextTaskCQ
            End If
            
            ' Récupérer le métier
            If onlyDedicatedCQ Then
                ' CQ dédiée : chercher la tâche d'origine
                metier = GetQualityTaskMetierFromOrigin(t)
                If Len(metier) = 0 Then metier = "CQ"  ' Fallback
                cqDediees = cqDediees + 1
            Else
                ' CQ sur tâche normale : utiliser Text4 directement
                metier = text4Val
                cqNormales = cqNormales + 1
            End If
            
            If Len(metier) = 0 Then GoTo NextTaskCQ
            
            ' Récupérer % Complete
            On Error Resume Next
            pct = t.PercentComplete
            If Err.Number <> 0 Then pct = 0
            On Error GoTo EH
            
            ' Clé : "Zone|Métier"
            key = zone & "|" & metier
            
            ' Accumulation
            If Not totalCountDict.Exists(key) Then
                totalCountDict(key) = 0
                completedCountDict(key) = 0
                sumPercentDict(key) = 0
            End If
            
            totalCountDict(key) = totalCountDict(key) + 1
            If pct = 100 Then completedCountDict(key) = completedCountDict(key) + 1
            sumPercentDict(key) = sumPercentDict(key) + pct
            
            ' Enregistrer zone et métier uniques
            If Not zonesOut.Exists(zone) Then zonesOut(zone) = True
            If Not metiersOut.Exists(metier) Then metiersOut(metier) = True
            
            totalCQ = totalCQ + 1
        End If
        
NextTaskCQ:
    Next t
    
    ' Logs récapitulatifs
    Debug.Print "=== RECAPITULATIF CQ ==="
    Debug.Print "Total tâches CQ: " & totalCQ
    Debug.Print "  - CQ sur tâche normale (Text4<>'CQ'): " & cqNormales
    Debug.Print "  - CQ dédiées (Text4='CQ'): " & cqDediees
    Debug.Print "Zones CQ: " & zonesOut.Count & " | Métiers CQ: " & metiersOut.Count
    
    ' Calcul final par clé
    Debug.Print "=== CALCUL FINAL ==="
    For Each k In totalCountDict.Keys
        Set dictItem = CreateObject("Scripting.Dictionary")
        dictItem("total") = totalCountDict(k)
        dictItem("completed") = completedCountDict(k)
        
        ' Moyenne des PercentComplete
        If totalCountDict(k) > 0 Then
            dictItem("avgPercent") = sumPercentDict(k) / totalCountDict(k)
        Else
            dictItem("avgPercent") = 0
        End If
        
        Set data(k) = dictItem
        
        Debug.Print k & ": Total=" & dictItem("total") & ", Terminés=" & dictItem("completed") & ", Moy=" & Format(dictItem("avgPercent"), "0.0") & "%"
    Next k
    
    Debug.Print "=== FIN ExtractQualityData ===" & vbCrLf
    
    Set ExtractQualityData = data
    Exit Function
    
EH:
    Debug.Print "ERREUR dans ExtractQualityData: " & Err.Number & " - " & Err.Description
    Set ExtractQualityData = data
End Function

Private Function GetQualityTaskMetierFromOrigin(ByVal tCQ As Task) As String
    ' Récupère le métier d'une tâche CQ dédiée en cherchant la tâche d'origine
    ' Utilisé quand Text4 = "CQ"
    
    Dim metier As String
    Dim nomCQ As String
    Dim nomOrigine As String
    Dim tOrigine As Task
    
    On Error GoTo EH
    
    nomCQ = tCQ.Name
    Debug.Print "  Tâche CQ dédiée [" & nomCQ & "] - Recherche tâche origine..."
    
    ' Vérifier si le nom commence par "Contrôle Qualité - "
    If InStr(1, nomCQ, "Contrôle Qualité - ", vbTextCompare) = 1 Then
        ' Extraire le nom après " - "
        nomOrigine = Mid(nomCQ, Len("Contrôle Qualité - ") + 1)
        
        ' Chercher la tâche avec ce nom
        For Each tOrigine In ActiveProject.Tasks
            If Not tOrigine Is Nothing And Not tOrigine.Summary Then
                On Error Resume Next
                If Trim(tOrigine.Name) = nomOrigine Then
                    metier = Trim(CStr(tOrigine.Text4))
                    If Err.Number = 0 And Len(metier) > 0 And UCase(metier) <> "CQ" Then
                        Debug.Print "    -> Tâche origine trouvée: [" & nomOrigine & "] avec Métier=" & metier
                        GetQualityTaskMetierFromOrigin = metier
                        Exit Function
                    End If
                End If
                On Error GoTo EH
            End If
        Next tOrigine
    End If
    
    ' Si pas trouvé, retourner chaîne vide
    Debug.Print "    -> Tâche origine NON trouvée"
    GetQualityTaskMetierFromOrigin = ""
    Exit Function
    
EH:
    Debug.Print "ERREUR dans GetQualityTaskMetierFromOrigin: " & Err.Number & " - " & Err.Description
    GetQualityTaskMetierFromOrigin = ""
End Function

Private Function GetQualityTaskMetier(ByVal tCQ As Task, ByRef cqNormalesCount As Long, ByRef cqDedieesCount As Long) As String
    ' DEPRECATED - Garder pour compatibilité avec traçabilité
    ' Récupère le métier d'une tâche CQ
    ' CAS 1 : Si Text4 <> "CQ", retourne Text4 directement (CQ sur tâche normale)
    ' CAS 2 : Si Text4 = "CQ", cherche la tâche d'origine (tâche CQ dédiée)
    
    Dim metier As String
    Dim text4Val As String
    
    On Error GoTo EH
    
    ' Récupérer Text4 de la tâche CQ
    On Error Resume Next
    text4Val = Trim(CStr(tCQ.Text4))
    If Err.Number <> 0 Then text4Val = ""
    On Error GoTo EH
    
    ' CAS 1 : Text4 <> "CQ" => CQ sur tâche normale
    If UCase(text4Val) <> "CQ" And Len(text4Val) > 0 Then
        metier = text4Val
        cqNormalesCount = cqNormalesCount + 1
        Debug.Print "  Tâche [" & tCQ.Name & "] - CAS 1 - Zone=" & tCQ.Text2 & " | Métier=" & metier & " | Pct=" & tCQ.PercentComplete & "%"
        GetQualityTaskMetier = metier
        Exit Function
    End If
    
    ' CAS 2 : Text4 = "CQ" => Tâche CQ dédiée
    metier = GetQualityTaskMetierFromOrigin(tCQ)
    If Len(metier) = 0 Then metier = "CQ"
    
    cqDedieesCount = cqDedieesCount + 1
    GetQualityTaskMetier = metier
    Exit Function
    
EH:
    Debug.Print "ERREUR dans GetQualityTaskMetier: " & Err.Number & " - " & Err.Description
    GetQualityTaskMetier = "CQ"
End Function

Private Sub CreateQualityTable(ByVal doc As Object, ByVal onlyDedicatedCQ As Boolean)
    ' Crée un tableau Word récapitulatif des CQ par (Zone, Métier)
    ' onlyDedicatedCQ: True = seulement CQ dédiées, False = seulement CQ sur tâches normales
    ' Colonnes : Zone | Métier | Nb Total | Nb Terminés | % Moyen
    
    Dim data As Object
    Dim zones As Object
    Dim metiers As Object
    Dim tbl As Object
    Dim rng As Object
    Dim sortedKeys() As String
    Dim numKeys As Long
    Dim i As Long, j As Long
    Dim key As Variant
    Dim dictItem As Object
    Dim zone As String
    Dim metier As String
    Dim nbTotal As Long
    Dim nbCompleted As Long
    Dim avgPercent As Double
    Dim colorRGB As Long
    
    On Error GoTo EH
    
    Debug.Print "=== DEBUT CreateQualityTable (onlyDedicatedCQ=" & onlyDedicatedCQ & ") ==="
    
    ' Extraire les données
    Set data = ExtractQualityData(onlyDedicatedCQ, zones, metiers)
    
    ' Vérifier qu'on a des données
    If data Is Nothing Or data.Count = 0 Then
        AddParagraph doc, "[Aucune tâche CQ détectée]", WD_ALIGN_LEFT
        Debug.Print "Aucune donnée CQ disponible"
        Exit Sub
    End If
    
    Debug.Print "Nombre de lignes à créer: " & data.Count
    
    ' Trier les clés par Zone puis Métier
    numKeys = data.Count
    ReDim sortedKeys(numKeys - 1)
    i = 0
    For Each key In data.Keys
        sortedKeys(i) = CStr(key)
        i = i + 1
    Next key
    
    ' Tri simple (bubble sort)
    Dim temp As String
    For i = 0 To numKeys - 2
        For j = i + 1 To numKeys - 1
            If sortedKeys(i) > sortedKeys(j) Then
                temp = sortedKeys(i)
                sortedKeys(i) = sortedKeys(j)
                sortedKeys(j) = temp
            End If
        Next j
    Next i
    
    ' Positionner le curseur
    Set rng = doc.Range(doc.Content.End - 1)
    
    ' Créer le tableau (nombre de lignes = données + 1 pour en-tête)
    Set tbl = doc.Tables.Add(rng, numKeys + 1, 5)
    
    ' Style du tableau
    On Error Resume Next
    tbl.Style = "Grille du tableau moyenne 2"
    tbl.AutoFitBehavior 2  ' wdAutoFitContent = 2
    On Error GoTo EH
    
    ' En-têtes (ligne 1)
    tbl.Cell(1, 1).Range.Text = "Zone"
    tbl.Cell(1, 1).Range.Bold = True
    tbl.Cell(1, 1).Shading.BackgroundPatternColor = RGB(200, 200, 200)
    tbl.Cell(1, 1).Range.ParagraphFormat.alignment = WD_ALIGN_CENTER
    
    tbl.Cell(1, 2).Range.Text = "Métier"
    tbl.Cell(1, 2).Range.Bold = True
    tbl.Cell(1, 2).Shading.BackgroundPatternColor = RGB(200, 200, 200)
    tbl.Cell(1, 2).Range.ParagraphFormat.alignment = WD_ALIGN_CENTER
    
    tbl.Cell(1, 3).Range.Text = "Nb CQ Total"
    tbl.Cell(1, 3).Range.Bold = True
    tbl.Cell(1, 3).Shading.BackgroundPatternColor = RGB(200, 200, 200)
    tbl.Cell(1, 3).Range.ParagraphFormat.alignment = WD_ALIGN_CENTER
    
    tbl.Cell(1, 4).Range.Text = "Nb CQ Terminés"
    tbl.Cell(1, 4).Range.Bold = True
    tbl.Cell(1, 4).Shading.BackgroundPatternColor = RGB(200, 200, 200)
    tbl.Cell(1, 4).Range.ParagraphFormat.alignment = WD_ALIGN_CENTER
    
    tbl.Cell(1, 5).Range.Text = "% Complet Moyen"
    tbl.Cell(1, 5).Range.Bold = True
    tbl.Cell(1, 5).Shading.BackgroundPatternColor = RGB(200, 200, 200)
    tbl.Cell(1, 5).Range.ParagraphFormat.alignment = WD_ALIGN_CENTER
    
    Debug.Print "En-têtes créés"
    
    ' Remplir les données (lignes suivantes)
    For i = 0 To numKeys - 1
        key = sortedKeys(i)
        Set dictItem = data(key)
        
        ' Extraire zone et métier de la clé "Zone|Métier"
        zone = Left(key, InStr(1, key, "|") - 1)
        metier = Mid(key, InStr(1, key, "|") + 1)
        
        nbTotal = dictItem("total")
        nbCompleted = dictItem("completed")
        avgPercent = dictItem("avgPercent")
        
        ' Colonne 1 : Zone (en gras)
        tbl.Cell(i + 2, 1).Range.Text = zone
        tbl.Cell(i + 2, 1).Range.Bold = True
        tbl.Cell(i + 2, 1).Range.ParagraphFormat.alignment = WD_ALIGN_CENTER
        
        ' Colonne 2 : Métier
        tbl.Cell(i + 2, 2).Range.Text = metier
        tbl.Cell(i + 2, 2).Range.ParagraphFormat.alignment = WD_ALIGN_LEFT
        
        ' Colonne 3 : Nb Total (centré)
        tbl.Cell(i + 2, 3).Range.Text = CStr(nbTotal)
        tbl.Cell(i + 2, 3).Range.ParagraphFormat.alignment = WD_ALIGN_CENTER
        
        ' Colonne 4 : Nb Terminés (centré)
        tbl.Cell(i + 2, 4).Range.Text = CStr(nbCompleted)
        tbl.Cell(i + 2, 4).Range.ParagraphFormat.alignment = WD_ALIGN_CENTER
        
        ' Colonne 5 : % Moyen avec 1 décimale + symbole % + coloration
        tbl.Cell(i + 2, 5).Range.Text = Format(avgPercent, "0.0") & "%"
        tbl.Cell(i + 2, 5).Range.ParagraphFormat.alignment = WD_ALIGN_CENTER
        
        ' Coloration selon la valeur
        On Error Resume Next
        If avgPercent >= 80 Then
            colorRGB = RGB(200, 255, 200)  ' Vert clair
        ElseIf avgPercent >= 50 Then
            colorRGB = RGB(255, 255, 200)  ' Jaune clair
        Else
            colorRGB = RGB(255, 200, 200)  ' Rouge clair
        End If
        tbl.Cell(i + 2, 5).Shading.BackgroundPatternColor = colorRGB
        On Error GoTo EH
        
        Debug.Print "  Ligne " & (i + 2) & ": " & zone & " | " & metier & " | Total=" & nbTotal & " | Terminés=" & nbCompleted & " | Moy=" & Format(avgPercent, "0.0") & "%"
    Next i
    
    ' Saut de ligne après le tableau
    rng.InsertAfter vbCrLf
    rng.Collapse WD_COLLAPSE_END
    
    Debug.Print "=== TABLEAU QUALITE CREE AVEC SUCCES ===" & vbCrLf
    
    Exit Sub
    
EH:
    Debug.Print "ERREUR CreateQualityTable: " & Err.Number & " - " & Err.Description
    On Error Resume Next
    AddParagraph doc, "[Erreur création tableau CQ: " & Err.Description & "]", WD_ALIGN_LEFT
    On Error GoTo 0
End Sub

Private Sub CreateQualityChart(ByVal doc As Object, ByVal onlyDedicatedCQ As Boolean)
    ' Crée un graphique en colonnes groupées des CQ par (Zone, Métier)
    ' onlyDedicatedCQ: True = seulement CQ dédiées, False = seulement CQ sur tâches normales
    ' Réutilise AddMultiSeriesChart existante
    
    Dim dataRaw As Object
    Dim dataChart As Object
    Dim zones As Object
    Dim metiers As Object
    Dim k As Variant
    Dim dictItem As Object
    Dim avgPercent As Double
    
    On Error GoTo EH
    
    Debug.Print "=== DEBUT CreateQualityChart (onlyDedicatedCQ=" & onlyDedicatedCQ & ") ==="
    
    ' Extraire les données
    Set dataRaw = ExtractQualityData(onlyDedicatedCQ, zones, metiers)
    
    ' Vérifier qu'on a des données
    If dataRaw Is Nothing Or dataRaw.Count = 0 Then
        AddParagraph doc, "[Aucune donnée CQ disponible pour le graphique]", WD_ALIGN_LEFT
        Debug.Print "Aucune donnée CQ pour graphique"
        Exit Sub
    End If
    
    ' Convertir le format de données : { "Zone|Métier": {total, completed, avgPercent} } -> { "Zone|Métier": avgPercent }
    Set dataChart = CreateObject("Scripting.Dictionary")
    For Each k In dataRaw.Keys
        Set dictItem = dataRaw(k)
        avgPercent = dictItem("avgPercent")
        dataChart(k) = avgPercent
    Next k
    
    Debug.Print "Données converties pour graphique: " & dataChart.Count & " entrées"
    
    ' Créer le graphique (ou tableau si échec)
    AddMultiSeriesChart doc, dataChart, zones, metiers, "Avancement des contrôles qualité par zone et métier"
    
    Debug.Print "=== FIN CreateQualityChart ===" & vbCrLf
    
    Exit Sub
    
EH:
    Debug.Print "ERREUR CreateQualityChart: " & Err.Number & " - " & Err.Description
    On Error Resume Next
    AddParagraph doc, "[Erreur création graphique CQ: " & Err.Description & "]", WD_ALIGN_LEFT
    On Error GoTo 0
End Sub

' =========================
' PROJECT DATA (MS Project)
' =========================

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

' =============================================================================
' SECTION TRAÇABILITÉ - Export données MS Project pour validation
' =============================================================================
' Cette section génère un fichier .txt détaillé qui permet de tracer l'origine
' de chaque donnée affichée dans les graphiques/tableaux du rapport Word.
' 
' Point d'entrée : ExportProjectDataTrace(outFolder)
' Appelé automatiquement depuis BuildWeeklyReport()
' 
' Fonctions incluses :
' - ExportProjectDataTrace : orchestrateur principal
' - TraceExportRawTaskList : liste brute de toutes les tâches
' - TraceExportProgressDetails : détail des calculs Section 2 (avancement)
' - TraceExportQualityDetails : détail des calculs Section 3 (CQ)
' =============================================================================

' =========================
' TRAÇABILITÉ - ORCHESTRATEUR PRINCIPAL
' =========================
Private Sub ExportProjectDataTrace(ByVal outFolder As String)
    ' Génère un fichier .txt contenant :
    ' - Liste brute de toutes les tâches MS Project
    ' - Détail des calculs pour chaque graphique Section 2 (4 graphiques)
    ' - Détail des calculs pour Section 3 (Contrôles Qualité)
    
    Dim txtPath As String
    Dim fso As Object
    Dim txtFile As Object
    
    On Error GoTo EH
    
    txtPath = outFolder & "\Rapport_Data_Trace_" & GetReportDateTimeStamp() & ".txt"
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set txtFile = fso.CreateTextFile(txtPath, True)
    
    ' En-tête du fichier
    txtFile.WriteLine String(80, "=")
    txtFile.WriteLine "TRAÇABILITÉ DES DONNÉES - MS PROJECT → RAPPORT PREVENCHERES"
    txtFile.WriteLine String(80, "=")
    txtFile.WriteLine "Date génération : " & Now
    txtFile.WriteLine "Projet MS Project : " & ActiveProject.Name
    txtFile.WriteLine "Nombre total de tâches : " & ActiveProject.Tasks.Count
    txtFile.WriteLine String(80, "=")
    txtFile.WriteLine ""
    txtFile.WriteLine ""
    
    ' PARTIE 1 : Liste brute de toutes les tâches
    txtFile.WriteLine String(80, "=")
    txtFile.WriteLine "PARTIE 1 : LISTE BRUTE DE TOUTES LES TÂCHES"
    txtFile.WriteLine String(80, "=")
    txtFile.WriteLine ""
    TraceExportRawTaskList txtFile
    txtFile.WriteLine ""
    txtFile.WriteLine ""
    
    ' PARTIE 2 : Section 2 - Graphique 2.1 (Zone × Métier, % tâches)
    txtFile.WriteLine String(80, "=")
    txtFile.WriteLine "PARTIE 2 : SECTION 2 - GRAPHIQUE 2.1"
    txtFile.WriteLine "Avancement par Zone et Métier (% tâches)"
    txtFile.WriteLine String(80, "=")
    txtFile.WriteLine ""
    TraceExportProgressDetails txtFile, "Zone", True
    txtFile.WriteLine ""
    txtFile.WriteLine ""
    
    ' PARTIE 3 : Section 2 - Graphique 2.2 (Zone × Métier, % ressources)
    txtFile.WriteLine String(80, "=")
    txtFile.WriteLine "PARTIE 3 : SECTION 2 - GRAPHIQUE 2.2"
    txtFile.WriteLine "Avancement par Zone et Métier (% ressources)"
    txtFile.WriteLine String(80, "=")
    txtFile.WriteLine ""
    TraceExportProgressDetails txtFile, "Zone", False
    txtFile.WriteLine ""
    txtFile.WriteLine ""
    
    ' PARTIE 4 : Section 2 - Graphique 2.3 (SousZone × Métier, % tâches)
    txtFile.WriteLine String(80, "=")
    txtFile.WriteLine "PARTIE 4 : SECTION 2 - GRAPHIQUE 2.3"
    txtFile.WriteLine "Avancement par Sous-Zone et Métier (% tâches)"
    txtFile.WriteLine String(80, "=")
    txtFile.WriteLine ""
    TraceExportProgressDetails txtFile, "SousZone", True
    txtFile.WriteLine ""
    txtFile.WriteLine ""
    
    ' PARTIE 5 : Section 2 - Graphique 2.4 (SousZone × Métier, % ressources)
    txtFile.WriteLine String(80, "=")
    txtFile.WriteLine "PARTIE 5 : SECTION 2 - GRAPHIQUE 2.4"
    txtFile.WriteLine "Avancement par Sous-Zone et Métier (% ressources)"
    txtFile.WriteLine String(80, "=")
    txtFile.WriteLine ""
    TraceExportProgressDetails txtFile, "SousZone", False
    txtFile.WriteLine ""
    txtFile.WriteLine ""
    
    ' PARTIE 6 : Section 3 - Contrôles Qualité sur tâches normales
    txtFile.WriteLine String(80, "=")
    txtFile.WriteLine "PARTIE 6A : SECTION 3 - CONTRÔLES QUALITÉ SUR TÂCHES NORMALES"
    txtFile.WriteLine "CQ avec Text4 <> 'CQ'"
    txtFile.WriteLine String(80, "=")
    txtFile.WriteLine ""
    TraceExportQualityDetails txtFile, False  ' False = CQ normales
    txtFile.WriteLine ""
    
    ' PARTIE 7 : Section 3 - Contrôles Qualité dédiées
    txtFile.WriteLine String(80, "=")
    txtFile.WriteLine "PARTIE 6B : SECTION 3 - CONTRÔLES QUALITÉ DÉDIÉES"
    txtFile.WriteLine "Tâches CQ avec Text4 = 'CQ'"
    txtFile.WriteLine String(80, "=")
    txtFile.WriteLine ""
    TraceExportQualityDetails txtFile, True  ' True = CQ dédiées
    txtFile.WriteLine ""
    
    ' Footer
    txtFile.WriteLine ""
    txtFile.WriteLine String(80, "=")
    txtFile.WriteLine "FIN DU FICHIER DE TRAÇABILITÉ"
    txtFile.WriteLine "Fichier : " & txtPath
    txtFile.WriteLine String(80, "=")
    
    txtFile.Close
    Set txtFile = Nothing
    Set fso = Nothing
    
    Debug.Print "✓ Fichier de traçabilité créé : " & txtPath
    
    Exit Sub
    
EH:
    Debug.Print "ERREUR ExportProjectDataTrace : " & Err.Number & " - " & Err.Description
    On Error Resume Next
    If Not txtFile Is Nothing Then txtFile.Close
    On Error GoTo 0
End Sub

' =========================
' TRAÇABILITÉ - LISTE BRUTE
' =========================
Private Sub TraceExportRawTaskList(ByRef txtFile As Object)
    ' Exporte la liste brute de toutes les tâches avec leurs propriétés principales
    ' Format : [ID] | Nom | Zone | SousZone | Métier | Work(h) | ActualWork(h) | %Complete | Ressources | Summary
    
    Dim t As Task
    Dim a As Object
    Dim resourcesList As String
    Dim ligne As String
    Dim taskCount As Long
    
    On Error Resume Next
    
    txtFile.WriteLine "Format des colonnes :"
    txtFile.WriteLine "[ID] | Nom | Zone | SousZone | Métier | Work(h) | ActualWork(h) | %Complete | Ressources | Summary"
    txtFile.WriteLine String(80, "-")
    txtFile.WriteLine ""
    
    For Each t In ActiveProject.Tasks
        If Not t Is Nothing Then
            taskCount = taskCount + 1
            
            ' Récupérer les ressources affectées
            resourcesList = ""
            For Each a In t.Assignments
                If Not a Is Nothing And Not a.Resource Is Nothing Then
                    If Len(resourcesList) > 0 Then resourcesList = resourcesList & ", "
                    resourcesList = resourcesList & a.Resource.Name
                End If
            Next a
            If Len(resourcesList) = 0 Then resourcesList = "[Aucune]"
            
            ' Construire la ligne
            ligne = "[" & t.ID & "] | "
            ligne = ligne & t.Name & " | "
            ligne = ligne & Trim(CStr(t.Text2)) & " | "
            ligne = ligne & Trim(CStr(t.Text3)) & " | "
            ligne = ligne & Trim(CStr(t.Text4)) & " | "
            ligne = ligne & Format(t.Work / 60, "0.00") & " | "
            ligne = ligne & Format(t.ActualWork / 60, "0.00") & " | "
            ligne = ligne & Format(t.PercentComplete, "0.0") & "% | "
            ligne = ligne & resourcesList & " | "
            ligne = ligne & IIf(t.Summary, "OUI", "NON")
            
            txtFile.WriteLine ligne
        End If
    Next t
    
    txtFile.WriteLine ""
    txtFile.WriteLine "Total tâches listées : " & taskCount
    
    On Error GoTo 0
End Sub

' =========================
' TRAÇABILITÉ - DÉTAIL AVANCEMENT (SECTION 2)
' =========================
Private Sub TraceExportProgressDetails(ByRef txtFile As Object, ByVal groupBy As String, ByVal useTaskPercent As Boolean)
    ' Exporte les détails de calcul pour un graphique d'avancement Section 2
    ' Montre pour chaque combinaison (Zone|Métier) :
    ' - Quelles tâches contribuent au calcul
    ' - Le détail du calcul étape par étape
    ' - Le résultat final qui apparaît dans le graphique
    
    Dim t As Task
    Dim zone As String
    Dim metier As String
    Dim key As String
    Dim tasksDict As Object  ' key -> Collection de lignes de tâches
    Dim percentDict As Object
    Dim percentWorkDict As Object
    Dim countDict As Object
    Dim finalPercent As Double
    Dim k As Variant
    Dim taskList As Collection
    Dim taskInfo As Variant
    Dim i As Long
    
    On Error Resume Next
    
    Set tasksDict = CreateObject("Scripting.Dictionary")
    Set percentDict = CreateObject("Scripting.Dictionary")
    Set percentWorkDict = CreateObject("Scripting.Dictionary")
    Set countDict = CreateObject("Scripting.Dictionary")
    
    txtFile.WriteLine "Type de calcul : " & IIf(useTaskPercent, "Moyenne des % Achevé (PercentComplete)", "Moyenne des % Travail Achevé (PercentWorkComplete)")
    txtFile.WriteLine "Groupement : " & groupBy & IIf(groupBy = "Zone", " (Text2)", " (Text3)")
    txtFile.WriteLine ""
    
    ' Collecte des tâches par (Zone|Métier)
    For Each t In ActiveProject.Tasks
        If Not t Is Nothing And Not t.Summary Then
            If t.Work > 0 Then
                ' Récupérer Zone ou SousZone
                If groupBy = "Zone" Then
                    zone = Trim(CStr(t.Text2))
                Else
                    zone = Trim(CStr(t.Text3))
                End If
                
                metier = Trim(CStr(t.Text4))
                
                If Len(zone) > 0 And Len(metier) > 0 Then
                    key = zone & "|" & metier
                    
                    ' Créer la collection si nécessaire
                    If Not tasksDict.Exists(key) Then
                        Set tasksDict(key) = New Collection
                        percentDict(key) = 0
                        percentWorkDict(key) = 0
                        countDict(key) = 0
                    End If
                    
                    ' Ajouter les informations de la tâche
                    taskInfo = "[" & t.ID & "] " & t.Name & " : " & _
                               "% Achevé=" & Format(t.PercentComplete, "0.0") & "% | " & _
                               "% Travail Achevé=" & Format(t.PercentWorkComplete, "0.0") & "%"
                    
                    tasksDict(key).Add taskInfo
                    
                    ' Accumuler
                    percentDict(key) = percentDict(key) + t.PercentComplete
                    percentWorkDict(key) = percentWorkDict(key) + t.PercentWorkComplete
                    countDict(key) = countDict(key) + 1
                End If
            End If
        End If
    Next t
    
    ' Afficher les résultats par clé
    If tasksDict.Count = 0 Then
        txtFile.WriteLine "[Aucune donnée disponible pour ce graphique]"
        txtFile.WriteLine ""
        On Error GoTo 0
        Exit Sub
    End If
    
    For Each k In tasksDict.Keys
        ' Calcul du résultat final
        If useTaskPercent Then
            If countDict(k) > 0 Then
                finalPercent = percentDict(k) / countDict(k)
            Else
                finalPercent = 0
            End If
        Else
            If countDict(k) > 0 Then
                finalPercent = percentWorkDict(k) / countDict(k)
            Else
                finalPercent = 0
            End If
        End If
        
        ' Afficher le header
        txtFile.WriteLine String(80, "-")
        txtFile.WriteLine "📊 " & Replace(k, "|", " | ") & " => " & Format(finalPercent, "0.0") & "%"
        txtFile.WriteLine "   Nombre de tâches : " & countDict(k)
        txtFile.WriteLine "   Détail des tâches :"
        
        ' Lister les tâches
        Set taskList = tasksDict(k)
        i = 1
        For Each taskInfo In taskList
            If i = taskList.Count Then
                txtFile.WriteLine "   └─ " & taskInfo
            Else
                txtFile.WriteLine "   ├─ " & taskInfo
            End If
            i = i + 1
        Next
        
        ' Afficher le calcul
        txtFile.WriteLine ""
        If useTaskPercent Then
            txtFile.WriteLine "   Calcul (moyenne % Achevé) :"
            txtFile.WriteLine "   = " & Format(percentDict(k), "0.0") & " / " & countDict(k)
            txtFile.WriteLine "   = " & Format(finalPercent, "0.0") & "%"
        Else
            txtFile.WriteLine "   Calcul (moyenne % Travail Achevé) :"
            txtFile.WriteLine "   = " & Format(percentWorkDict(k), "0.0") & " / " & countDict(k)
            txtFile.WriteLine "   = " & Format(finalPercent, "0.0") & "%"
        End If
        
        txtFile.WriteLine ""
    Next k
    
    txtFile.WriteLine ""
    txtFile.WriteLine "Total combinaisons (Zone|Métier) : " & tasksDict.Count
    
    On Error GoTo 0
End Sub

' =========================
' TRAÇABILITÉ - DÉTAIL CONTRÔLES QUALITÉ (SECTION 3)
' =========================
Private Sub TraceExportQualityDetails(ByRef txtFile As Object, ByVal onlyDedicatedCQ As Boolean)
    ' Exporte les détails des contrôles qualité (tâches avec ressource CQ)
    ' onlyDedicatedCQ: True = seulement CQ dédiées (Text4="CQ"), False = seulement CQ sur tâches normales
    ' Montre pour chaque combinaison (Zone|Métier) :
    ' - Quelles tâches CQ contribuent au calcul
    ' - Le nombre total et le nombre terminé
    ' - Le % moyen qui apparaît dans le tableau/graphique
    
    Dim t As Task
    Dim a As Object
    Dim hasCQ As Boolean
    Dim zone As String
    Dim metier As String
    Dim text4Val As String
    Dim key As String
    Dim tasksDict As Object
    Dim totalDict As Object
    Dim completedDict As Object
    Dim sumPercentDict As Object
    Dim taskList As Collection
    Dim taskInfo As Variant
    Dim k As Variant
    Dim avgPercent As Double
    Dim i As Long
    Dim cqCount As Long
    
    On Error Resume Next
    
    Set tasksDict = CreateObject("Scripting.Dictionary")
    Set totalDict = CreateObject("Scripting.Dictionary")
    Set completedDict = CreateObject("Scripting.Dictionary")
    Set sumPercentDict = CreateObject("Scripting.Dictionary")
    
    txtFile.WriteLine "Filtre : Tâches avec ressource 'CQ' affectée"
    If onlyDedicatedCQ Then
        txtFile.WriteLine "Type : CQ DÉDIÉES (Text4 = 'CQ')"
    Else
        txtFile.WriteLine "Type : CQ SUR TÂCHES NORMALES (Text4 <> 'CQ')"
    End If
    txtFile.WriteLine ""
    
    ' Collecte des tâches CQ
    For Each t In ActiveProject.Tasks
        If Not t Is Nothing And Not t.Summary Then
            ' Vérifier si ressource CQ affectée
            hasCQ = False
            For Each a In t.Assignments
                If Not a Is Nothing And Not a.Resource Is Nothing Then
                    If UCase(Trim(a.Resource.Name)) = "CQ" Then
                        hasCQ = True
                        Exit For
                    End If
                End If
            Next a
            
            If hasCQ Then
                zone = Trim(CStr(t.Text2))
                text4Val = Trim(CStr(t.Text4))
                
                ' Filtrer selon le type demandé
                If onlyDedicatedCQ Then
                    ' On veut seulement les CQ dédiées (Text4 = "CQ")
                    If UCase(text4Val) <> "CQ" Then GoTo NextCQTrace
                    metier = GetQualityTaskMetierFromOrigin(t)
                    If Len(metier) = 0 Then metier = "CQ"
                Else
                    ' On veut seulement les CQ sur tâches normales (Text4 <> "CQ")
                    If UCase(text4Val) = "CQ" Or Len(text4Val) = 0 Then GoTo NextCQTrace
                    metier = text4Val
                End If
                
                If Len(zone) > 0 And Len(metier) > 0 Then
                    key = zone & "|" & metier
                    
                    If Not tasksDict.Exists(key) Then
                        Set tasksDict(key) = New Collection
                        totalDict(key) = 0
                        completedDict(key) = 0
                        sumPercentDict(key) = 0
                    End If
                    
                    taskInfo = "[" & t.ID & "] " & t.Name & " : " & Format(t.PercentComplete, "0.0") & "%"
                    If t.PercentComplete = 100 Then taskInfo = taskInfo & " ✓"
                    
                    tasksDict(key).Add taskInfo
                    
                    totalDict(key) = totalDict(key) + 1
                    If t.PercentComplete = 100 Then completedDict(key) = completedDict(key) + 1
                    sumPercentDict(key) = sumPercentDict(key) + t.PercentComplete
                    
                    cqCount = cqCount + 1
                End If
            End If
NextCQTrace:
        End If
    Next t
    
    ' Afficher les résultats
    If tasksDict.Count = 0 Then
        txtFile.WriteLine "[Aucune tâche CQ de ce type détectée]"
        txtFile.WriteLine ""
        On Error GoTo 0
        Exit Sub
    End If
    
    txtFile.WriteLine "Nombre total de tâches CQ : " & cqCount
    txtFile.WriteLine ""
    
    For Each k In tasksDict.Keys
        avgPercent = 0
        If totalDict(k) > 0 Then avgPercent = sumPercentDict(k) / totalDict(k)
        
        txtFile.WriteLine String(80, "-")
        txtFile.WriteLine "📊 " & Replace(k, "|", " | ")
        txtFile.WriteLine "   Nb CQ Total : " & totalDict(k)
        txtFile.WriteLine "   Nb CQ Terminés (100%) : " & completedDict(k)
        txtFile.WriteLine "   % Complet Moyen : " & Format(avgPercent, "0.0") & "%"
        txtFile.WriteLine "   Détail des tâches CQ :"
        
        Set taskList = tasksDict(k)
        i = 1
        For Each taskInfo In taskList
            If i = taskList.Count Then
                txtFile.WriteLine "   └─ " & taskInfo
            Else
                txtFile.WriteLine "   ├─ " & taskInfo
            End If
            i = i + 1
        Next
        
        txtFile.WriteLine ""
        txtFile.WriteLine "   Calcul (% moyen) :"
        txtFile.WriteLine "   = " & Format(sumPercentDict(k), "0.0") & " / " & totalDict(k)
        txtFile.WriteLine "   = " & Format(avgPercent, "0.0") & "%"
        txtFile.WriteLine ""
    Next k
    
    txtFile.WriteLine ""
    txtFile.WriteLine "Total combinaisons (Zone|Métier) avec CQ : " & tasksDict.Count
    
    On Error GoTo 0
End Sub

