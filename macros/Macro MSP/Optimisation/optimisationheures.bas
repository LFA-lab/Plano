Option Explicit

' Types UDT pour la configuration et les KPI
Private Type Config
    seuilRelPct As Double
    seuilAbsH As Double
    seuilAbsQty As Double
    fieldPlan As String
    fieldReal As String
    fieldUnit As String
    exportsDir As String
End Type

Private Type TaskKpi
    ID As Long
    WBS As String
    Name As String
    IsParent As Boolean
    HasChild As Boolean
    OutlineLevel As Integer
    BaseH As Double   ' heures de base (planned / baseline)
    ActualH As Double ' heures consommées
    RemH As Double    ' heures restantes
    PVh As Double     ' Planned Value (en heures)
    EW As Double      ' Earned Work (en heures)
    EcartH As Double
    SPIh As Double
    CPIh As Double
    HeuresOptimisees As Double
End Type

' Variables globales pour le logging
Private logMessages As Object
Private logLevel As Integer

' Point d'entrée principal
Public Sub ExportOptimisationHeuresEtQuantites()
    Dim cfg As Config
    Dim xlApp As Object, xlWb As Object
    Dim projectName As String, exportPath As String
    Dim dtEtat As Date
    Dim lastS0Path As String, s0Data As Object
    
    ' Initialisation du logging
    Set logMessages = CreateObject("Scripting.Dictionary")
    logLevel = 0
    
    LogInfo "=== DÉBUT EXPORT OPTIMISATION ==="
    
    ' Chargement de la configuration
    If Not LoadOrPromptConfig(cfg) Then
        LogError "Configuration invalide, arrêt du traitement"
        Exit Sub
    End If
    
    ' Vérification baseline
    If Not EnsureBaselineOrWarn() Then
        LogWarn "Baseline manquante - KPI limités"
    End If
    
    ' Détermination de la date d'état (robuste si non définie)
    If IsDate(ActiveProject.StatusDate) And ActiveProject.StatusDate <> 0 Then
        dtEtat = ActiveProject.StatusDate
    Else
        ' si la date d'état n'est pas définie, on prend la date du jour (ou proposer un InputBox selon besoin)
        dtEtat = Date
    End If
    LogInfo "Date d'état utilisée : " & Format(dtEtat, "dd/mm/yyyy")
    
    ' Création du fichier Excel
    projectName = ActiveProject.Name
    If Right(projectName, 4) = ".mpp" Then
        projectName = Left(projectName, Len(projectName) - 4)
    End If
    
    exportPath = cfg.exportsDir & projectName & "_" & Format(Date, "yyyy") & "-" & Format(DatePart("ww", Date), "00") & ".xlsx"
    LogInfo "Export vers : " & exportPath
    
    ' Chargement du dernier snapshot S0
    lastS0Path = GetLastExportPath(projectName)
    If lastS0Path <> "" Then
        Set s0Data = LoadSnapshot(lastS0Path)
        LogInfo "S0 chargé depuis : " & lastS0Path
    Else
        Set s0Data = Nothing
        LogInfo "Aucun S0 trouvé - première exportation"
    End If
    
    ' Création de l'application Excel (Late Binding)
    On Error Resume Next
    Set xlApp = GetObject(, "Excel.Application")
    If xlApp Is Nothing Then
        Set xlApp = CreateObject("Excel.Application")
    End If
    On Error GoTo 0
    
    xlApp.Visible = False
    xlApp.ScreenUpdating = False
    xlApp.Calculation = -4105 ' xlCalculationManual
    
    ' Création du classeur
    Set xlWb = xlApp.Workbooks.Add
    
    ' Création des 7 onglets dans l'ordre
    Call CreateWorksheets(xlWb)
    
    ' Export des données vers chaque onglet
    Call WriteSheet_RESUME_DIRIGEANT(xlWb, cfg, dtEtat, s0Data)
    Call WriteSheet_RECAP_LOTS(xlWb, cfg, dtEtat, s0Data)
    Call WriteSheet_TACHES_PARENTS(xlWb, cfg, dtEtat, s0Data)
    Call WriteSheet_SOUS_TACHES_ENFANTS(xlWb, cfg, dtEtat, s0Data)
    Call WriteSheet_CONSOMMABLES(xlWb, cfg)
    Call WriteSheet_GUIDAGE(xlWb, lastS0Path)
    Call WriteSheet_LOG(xlWb)
    
    ' Sauvegarde et fermeture
    xlWb.SaveAs exportPath
    xlWb.Close
    
    xlApp.ScreenUpdating = True
    xlApp.Calculation = -4105 ' xlCalculationAutomatic
    xlApp.Quit
    
    LogInfo "=== EXPORT TERMINÉ ==="
    MsgBox "Export terminé avec succès !" & vbCrLf & exportPath
End Sub

' Chargement ou création de la configuration
Private Function LoadOrPromptConfig(ByRef cfg As Config) As Boolean
    Dim configPath As String
    Dim fso As Object, ts As Object
    Dim line As String, parts As Variant
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    configPath = fso.GetParentFolderName(ActiveProject.Path) & "\config.ini"
    
    ' Valeurs par défaut
    cfg.seuilRelPct = 3
    cfg.seuilAbsH = 2
    cfg.seuilAbsQty = 1
    cfg.fieldPlan = "Number1"
    cfg.fieldReal = "Number2"
    cfg.fieldUnit = "Text1"
    cfg.exportsDir = fso.GetParentFolderName(ActiveProject.Path) & "\Optimisation\Exports\"
    
    ' Lecture du fichier de config s'il existe
    If fso.FileExists(configPath) Then
        Set ts = fso.OpenTextFile(configPath, 1)
        Do While Not ts.AtEndOfStream
            line = Trim(ts.ReadLine)
            If InStr(line, "=") > 0 Then
                parts = Split(line, "=")
                Select Case Trim(parts(0))
                    Case "seuilRel%": cfg.seuilRelPct = Val(parts(1))
                    Case "seuilAbsH": cfg.seuilAbsH = Val(parts(1))
                    Case "seuilAbsQty": cfg.seuilAbsQty = Val(parts(1))
                    Case "fieldPlan": cfg.fieldPlan = Trim(parts(1))
                    Case "fieldReal": cfg.fieldReal = Trim(parts(1))
                    Case "fieldUnit": cfg.fieldUnit = Trim(parts(1))
                End Select
            End If
        Loop
        ts.Close
        LogInfo "Configuration chargée depuis " & configPath
    Else
        ' Création du fichier de config par défaut
        If Not fso.FolderExists(fso.GetParentFolderName(cfg.exportsDir)) Then
            fso.CreateFolder fso.GetParentFolderName(cfg.exportsDir)
        End If
        If Not fso.FolderExists(cfg.exportsDir) Then
            fso.CreateFolder cfg.exportsDir
        End If
        
        Set ts = fso.CreateTextFile(configPath, True)
        ts.WriteLine "seuilRel%=" & cfg.seuilRelPct
        ts.WriteLine "seuilAbsH=" & cfg.seuilAbsH
        ts.WriteLine "seuilAbsQty=" & cfg.seuilAbsQty
        ts.WriteLine "fieldPlan=" & cfg.fieldPlan
        ts.WriteLine "fieldReal=" & cfg.fieldReal
        ts.WriteLine "fieldUnit=" & cfg.fieldUnit
        ts.Close
        LogInfo "Configuration par défaut créée : " & configPath
    End If
    
    LoadOrPromptConfig = True
End Function

' Vérification de l'existence d'une baseline
Private Function EnsureBaselineOrWarn() As Boolean
    Dim t As Task
    Dim hasBaseline As Boolean
    
    hasBaseline = False
    For Each t In ActiveProject.Tasks
        If Not t Is Nothing And Not t.Summary Then
            If t.BaselineWork > 0 Then
                hasBaseline = True
                Exit For
            End If
        End If
    Next t
    
    EnsureBaselineOrWarn = hasBaseline
    If Not hasBaseline Then
        LogWarn "Aucune baseline détectée - les KPI SPI_h et CPI_h seront limités"
    End If
End Function

' Recherche du dernier export pour comparaison S0
Private Function GetLastExportPath(ByVal projectName As String) As String
    Dim fso As Object, folder As Object, files As Object, file As Object
    Dim exportDir As String
    Dim latestDate As Date, fileDate As Date
    Dim latestFile As String
    Dim pattern As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    exportDir = fso.GetParentFolderName(ActiveProject.Path) & "\Optimisation\Exports\"
    
    If Not fso.FolderExists(exportDir) Then
        GetLastExportPath = ""
        Exit Function
    End If
    
    Set folder = fso.GetFolder(exportDir)
    Set files = folder.Files
    
    pattern = projectName & "_"
    latestDate = CDate("01/01/1900")
    latestFile = ""
    
    For Each file In files
        If Left(file.Name, Len(pattern)) = pattern And Right(file.Name, 5) = ".xlsx" Then
            fileDate = file.DateLastModified
            If fileDate > latestDate Then
                latestDate = fileDate
                latestFile = file.Path
            End If
        End If
    Next file
    
    GetLastExportPath = latestFile
End Function

' Chargement des données snapshot S0
Private Function LoadSnapshot(ByVal path As String) As Object
    Dim xlApp As Object, xlWb As Object, xlWs As Object
    Dim dict As Object
    Dim i As Integer, lastRow As Integer
    Dim wbs As String, ecartH As Double
    
    Set dict = CreateObject("Scripting.Dictionary")
    
    On Error GoTo ErrorHandler
    
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = False
    Set xlWb = xlApp.Workbooks.Open(path)
    
    ' Chargement des données depuis SOUS_TACHES_ENFANTS
    Set xlWs = xlWb.Worksheets("SOUS_TACHES_ENFANTS")
    lastRow = xlWs.Cells(xlWs.Rows.Count, 1).End(-4162).Row ' xlUp
    
    For i = 2 To lastRow
        wbs = xlWs.Cells(i, 2).Value ' Colonne WBS
        ecartH = xlWs.Cells(i, 8).Value ' Colonne Écart_h
        If wbs <> "" Then
            dict(wbs) = ecartH
        End If
    Next i
    
    xlWb.Close False
    xlApp.Quit
    
    Set LoadSnapshot = dict
    Exit Function
    
ErrorHandler:
    If Not xlWb Is Nothing Then xlWb.Close False
    If Not xlApp Is Nothing Then xlApp.Quit
    Set LoadSnapshot = Nothing
    LogError "Erreur lors du chargement S0 : " & Err.Description
End Function

' Calcul des KPI pour une tâche
Private Sub ComputeTaskKpis(ByRef t As Task, ByRef k As TaskKpi, ByVal dtEtat As Date)
    k.ID = t.ID
    k.WBS = t.WBS
    k.Name = t.Name
    k.IsParent = t.Summary
    k.HasChild = (t.OutlineChildren.Count > 0)
    k.OutlineLevel = t.OutlineLevel
    
    ' Heures
    k.BaseH = t.BaselineWork / 60 ' Conversion minutes -> heures
    k.ActualH = t.ActualWork / 60
    k.RemH = t.RemainingWork / 60
    k.EcartH = k.ActualH - k.BaseH
    
    ' PV_h : BaselineWork des tâches avec BaselineFinish <= DateEtat
    If t.BaselineFinish <= dtEtat And t.BaselineFinish <> pjNA Then
        k.PVh = k.BaseH
    Else
        k.PVh = 0
    End If
    
    ' EW (Earned Work)
    k.EW = k.BaseH * (t.PercentComplete / 100)
    
    ' Indices SPI_h et CPI_h
    If k.PVh > 0 Then
        k.SPIh = k.EW / k.PVh
    Else
        k.SPIh = 0
    End If
    
    If k.ActualH > 0 Then
        k.CPIh = k.EW / k.ActualH
    Else
        k.CPIh = 0
    End If
End Sub

' Création des feuilles Excel
Private Sub CreateWorksheets(ByRef xlWb As Object)
    Dim sheetNames As Variant
    Dim i As Integer
    
    sheetNames = Array("RESUME_DIRIGEANT", "RECAP_LOTS", "TACHES_PARENTS", "SOUS_TACHES_ENFANTS", "CONSOMMABLES", "GUIDAGE", "LOG")
    
    ' Suppression des feuilles par défaut
    xlWb.Application.DisplayAlerts = False
    ' Supprimer toutes les feuilles sauf une, pour éviter l'erreur
    Do While xlWb.Worksheets.Count > 1
        xlWb.Worksheets(1).Delete
    Loop
    
    ' Effacer le contenu de la dernière feuille restante avant de créer les nouvelles
    xlWb.Worksheets(1).Cells.Clear
    xlWb.Application.DisplayAlerts = True
    
    ' Création des nouvelles feuilles
    For i = 0 To UBound(sheetNames)
        xlWb.Worksheets.Add.Name = sheetNames(i)
    Next i
End Sub

' Export vers RESUME_DIRIGEANT
Private Sub WriteSheet_RESUME_DIRIGEANT(ByRef xlWb As Object, ByRef cfg As Config, ByVal dtEtat As Date, ByVal s0Data As Object)
    Dim xlWs As Object
    Dim t As Task, k As TaskKpi
    Dim totalBase As Double, totalActual As Double, totalRem As Double
    Dim totalPV As Double, totalEW As Double
    Dim spiGlobal As Double, cpiGlobal As Double
    Dim heuresOptimisees As Double
    Dim row As Integer
    
    Set xlWs = xlWb.Worksheets("RESUME_DIRIGEANT")
    
    ' Calcul des totaux
    totalBase = 0#: totalActual = 0#: totalRem = 0#: totalPV = 0#: totalEW = 0#
    For Each t In ActiveProject.Tasks
        If Not t Is Nothing And Not t.Summary Then
            Call ComputeTaskKpis(t, k, dtEtat)
            totalBase = totalBase + k.BaseH
            totalActual = totalActual + k.ActualH
            totalRem = totalRem + k.RemH
            totalPV = totalPV + k.PVh
            totalEW = totalEW + k.EW
        End If
    Next t
    
    ' Calcul des indices globaux
    If totalPV > 0 Then spiGlobal = totalEW / totalPV
    If totalActual > 0 Then cpiGlobal = totalEW / totalActual
    
    ' Calcul des heures optimisées (logique à implémenter selon S0/S1)
    heuresOptimisees = 0 ' TODO: Calcul basé sur comparaison S0/S1
    
    ' Bloc KPI
    row = 1
    xlWs.Cells(row, 1) = "KPI PROJET"
    xlWs.Cells(row, 1).Font.Bold = True
    row = row + 1
    
    xlWs.Cells(row, 1) = "Heures prévues": xlWs.Cells(row, 2) = totalBase: row = row + 1
    xlWs.Cells(row, 1) = "Heures réelles": xlWs.Cells(row, 2) = totalActual: row = row + 1
    xlWs.Cells(row, 1) = "Heures restantes": xlWs.Cells(row, 2) = totalRem: row = row + 1
    xlWs.Cells(row, 1) = "Écart net (h)": xlWs.Cells(row, 2) = totalActual - totalBase: row = row + 1
    xlWs.Cells(row, 1) = "Heures optimisées (S0→S1)": xlWs.Cells(row, 2) = heuresOptimisees: row = row + 1
    xlWs.Cells(row, 1) = "SPI_h": xlWs.Cells(row, 2) = spiGlobal: row = row + 1
    xlWs.Cells(row, 1) = "CPI_h": xlWs.Cells(row, 2) = cpiGlobal: row = row + 1
    
    ' Formatage des colonnes
    xlWs.Range("B:B").NumberFormat = "0"
    xlWs.Columns("A:B").AutoFit
    
    LogInfo "RESUME_DIRIGEANT créé avec " & (row - 1) & " lignes de KPI"
End Sub

' Export vers RECAP_LOTS
Private Sub WriteSheet_RECAP_LOTS(ByRef xlWb As Object, ByRef cfg As Config, ByVal dtEtat As Date, ByVal s0Data As Object)
    Dim xlWs As Object
    Dim t As Task, k As TaskKpi
    Dim lots As Object
    Dim lotWbs As String, lotName As String
    Dim row As Integer
    
    Set xlWs = xlWb.Worksheets("RECAP_LOTS")
    Set lots = CreateObject("Scripting.Dictionary")
    
    ' En-têtes
    row = 1
    xlWs.Cells(row, 1) = "WBS"
    xlWs.Cells(row, 2) = "Lot / Phase"
    xlWs.Cells(row, 3) = "Base h"
    xlWs.Cells(row, 4) = "PV_h"
    xlWs.Cells(row, 5) = "EW"
    xlWs.Cells(row, 6) = "Actual"
    xlWs.Cells(row, 7) = "Rem."
    xlWs.Cells(row, 8) = "Écart_h"
    xlWs.Cells(row, 9) = "SPI_h"
    xlWs.Cells(row, 10) = "CPI_h"
    xlWs.Cells(row, 11) = "Heures_optimisées"
    
    xlWs.Range("1:1").Font.Bold = True
    xlWs.Range("1:1").AutoFilter
    
    ' Agrégation par lots (tâches de niveau 2)
    For Each t In ActiveProject.Tasks
        If Not t Is Nothing And t.OutlineLevel = 2 Then
            Call ComputeTaskKpis(t, k, dtEtat)
            lotWbs = k.WBS
            lotName = k.Name
            
            ' Agrégation des sous-tâches
            row = row + 1
            xlWs.Cells(row, 1) = lotWbs
            xlWs.Cells(row, 2) = lotName
            xlWs.Cells(row, 3) = k.BaseH
            xlWs.Cells(row, 4) = k.PVh
            xlWs.Cells(row, 5) = k.EW
            xlWs.Cells(row, 6) = k.ActualH
            xlWs.Cells(row, 7) = k.RemH
            xlWs.Cells(row, 8) = k.EcartH
            xlWs.Cells(row, 9) = k.SPIh
            xlWs.Cells(row, 10) = k.CPIh
            xlWs.Cells(row, 11) = 0 ' TODO: Calcul heures optimisées
        End If
    Next t
    
    ' Formatage
    xlWs.Range("C:H").NumberFormat = "0"
    xlWs.Range("I:J").NumberFormat = "0.00"
    xlWs.Range("K:K").NumberFormat = "0"
    xlWs.Columns.AutoFit
    
    LogInfo "RECAP_LOTS créé avec " & (row - 1) & " lots"
End Sub

' Export vers TACHES_PARENTS
Private Sub WriteSheet_TACHES_PARENTS(ByRef xlWb As Object, ByRef cfg As Config, ByVal dtEtat As Date, ByVal s0Data As Object)
    Dim xlWs As Object
    Dim t As Task, k As TaskKpi
    Dim row As Integer
    Dim drapeauH As String
    
    Set xlWs = xlWb.Worksheets("TACHES_PARENTS")
    
    ' En-têtes
    row = 1
    xlWs.Cells(row, 1) = "ID"
    xlWs.Cells(row, 2) = "WBS"
    xlWs.Cells(row, 3) = "Tâche parent"
    xlWs.Cells(row, 4) = "%C"
    xlWs.Cells(row, 5) = "Base"
    xlWs.Cells(row, 6) = "PV_h"
    xlWs.Cells(row, 7) = "EW"
    xlWs.Cells(row, 8) = "Actual"
    xlWs.Cells(row, 9) = "Rem."
    xlWs.Cells(row, 10) = "Écart_h"
    xlWs.Cells(row, 11) = "SPI_h"
    xlWs.Cells(row, 12) = "CPI_h"
    xlWs.Cells(row, 13) = "Drapeau_H"
    xlWs.Cells(row, 14) = "Heures_optimisées"
    xlWs.Cells(row, 15) = "A_un_enfant?"
    
    xlWs.Range("1:1").Font.Bold = True
    xlWs.Range("1:1").AutoFilter
    
    ' Export des tâches parentes
    For Each t In ActiveProject.Tasks
        If Not t Is Nothing And Not t.Summary And t.OutlineChildren.Count > 0 Then
            Call ComputeTaskKpis(t, k, dtEtat)
            
            ' Calcul du drapeau
            drapeauH = ""
            If k.BaseH > 0 Then
                If k.ActualH > k.BaseH * (1 + cfg.seuilRelPct / 100) And (k.ActualH - k.BaseH) >= cfg.seuilAbsH Then
                    drapeauH = "Dérive"
                ElseIf k.ActualH < k.BaseH * (1 - cfg.seuilRelPct / 100) And (k.BaseH - k.ActualH) >= cfg.seuilAbsH Then
                    drapeauH = "Économie"
                End If
            End If
            
            row = row + 1
            xlWs.Cells(row, 1) = k.ID
            xlWs.Cells(row, 2) = k.WBS
            xlWs.Cells(row, 3) = k.Name
            xlWs.Cells(row, 4) = t.PercentComplete
            xlWs.Cells(row, 5) = k.BaseH
            xlWs.Cells(row, 6) = k.PVh
            xlWs.Cells(row, 7) = k.EW
            xlWs.Cells(row, 8) = k.ActualH
            xlWs.Cells(row, 9) = k.RemH
            xlWs.Cells(row, 10) = k.EcartH
            xlWs.Cells(row, 11) = k.SPIh
            xlWs.Cells(row, 12) = k.CPIh
            xlWs.Cells(row, 13) = drapeauH
            xlWs.Cells(row, 14) = 0 ' TODO: Calcul heures optimisées
            xlWs.Cells(row, 15) = IIf(k.HasChild, "Oui", "Non")
        End If
    Next t
    
    ' Formatage
    xlWs.Range("D:D").NumberFormat = "0%"
    xlWs.Range("E:J").NumberFormat = "0"
    xlWs.Range("K:L").NumberFormat = "0.00"
    xlWs.Range("N:N").NumberFormat = "0"
    xlWs.Columns.AutoFit
    
    LogInfo "TACHES_PARENTS créé avec " & (row - 1) & " tâches parentes"
End Sub

' Export vers SOUS_TACHES_ENFANTS
Private Sub WriteSheet_SOUS_TACHES_ENFANTS(ByRef xlWb As Object, ByRef cfg As Config, ByVal dtEtat As Date, ByVal s0Data As Object)
    Dim xlWs As Object
    Dim t As Task, k As TaskKpi
    Dim row As Integer
    Dim parentWbs As String
    
    Set xlWs = xlWb.Worksheets("SOUS_TACHES_ENFANTS")
    
    ' En-têtes
    row = 1
    xlWs.Cells(row, 1) = "ID"
    xlWs.Cells(row, 2) = "WBS"
    xlWs.Cells(row, 3) = "Parent WBS"
    xlWs.Cells(row, 4) = "Sous-tâche"
    xlWs.Cells(row, 5) = "%C"
    xlWs.Cells(row, 6) = "Base"
    xlWs.Cells(row, 7) = "Actual"
    xlWs.Cells(row, 8) = "Écart_h"
    xlWs.Cells(row, 9) = "SPI_h"
    xlWs.Cells(row, 10) = "CPI_h"
    xlWs.Cells(row, 11) = "Heures_optimisées"
    xlWs.Cells(row, 12) = "Levier"
    xlWs.Cells(row, 13) = "Date_Action"
    
    xlWs.Range("1:1").Font.Bold = True
    xlWs.Range("1:1").AutoFilter
    
    ' Export des sous-tâches
    For Each t In ActiveProject.Tasks
        If Not t Is Nothing And Not t.Summary And t.OutlineChildren.Count = 0 And t.OutlineLevel > 1 Then
            Call ComputeTaskKpis(t, k, dtEtat)
            
            ' Recherche du parent WBS
            parentWbs = ""
            If Not t.OutlineParent Is Nothing Then
                parentWbs = t.OutlineParent.WBS
            End If
            
            row = row + 1
            xlWs.Cells(row, 1) = k.ID
            xlWs.Cells(row, 2) = k.WBS
            xlWs.Cells(row, 3) = parentWbs
            xlWs.Cells(row, 4) = k.Name
            xlWs.Cells(row, 5) = t.PercentComplete
            xlWs.Cells(row, 6) = k.BaseH
            xlWs.Cells(row, 7) = k.ActualH
            xlWs.Cells(row, 8) = k.EcartH
            xlWs.Cells(row, 9) = k.SPIh
            xlWs.Cells(row, 10) = k.CPIh
            xlWs.Cells(row, 11) = 0 ' TODO: Calcul heures optimisées selon S0/S1
            xlWs.Cells(row, 12) = "" ' À remplir depuis GUIDAGE
            xlWs.Cells(row, 13) = "" ' À remplir depuis GUIDAGE
        End If
    Next t
    
    ' Formatage
    xlWs.Range("E:E").NumberFormat = "0%"
    xlWs.Range("F:H").NumberFormat = "0"
    xlWs.Range("I:J").NumberFormat = "0.00"
    xlWs.Range("K:K").NumberFormat = "0"
    xlWs.Columns.AutoFit
    
    LogInfo "SOUS_TACHES_ENFANTS créé avec " & (row - 1) & " sous-tâches"
End Sub

' Export vers CONSOMMABLES
Private Sub WriteSheet_CONSOMMABLES(ByRef xlWb As Object, ByRef cfg As Config)
    Dim xlWs As Object
    Dim t As Task, r As Resource, a As Assignment
    Dim row As Integer
    Dim qtyPlan As Double, qtyReal As Double, ecartQty As Double, pctEcart As Double
    Dim drapeauQty As String
    Dim parentWbs As String, unite As String, article As String
    
    Set xlWs = xlWb.Worksheets("CONSOMMABLES")
    
    ' En-têtes
    row = 1
    xlWs.Cells(row, 1) = "Article (unité)"
    xlWs.Cells(row, 2) = "Parent WBS"
    xlWs.Cells(row, 3) = "Sous-tâche"
    xlWs.Cells(row, 4) = "Qty_plan"
    xlWs.Cells(row, 5) = "Qty_réel"
    xlWs.Cells(row, 6) = "Écart_qty"
    xlWs.Cells(row, 7) = "%"
    xlWs.Cells(row, 8) = "Drapeau_Qty"
    
    xlWs.Range("1:1").Font.Bold = True
    xlWs.Range("1:1").AutoFilter
    
    ' Parcours des ressources matérielles
    For Each r In ActiveProject.Resources
        If Not r Is Nothing And r.Type = 1 Then ' pjResourceTypeMaterial
            For Each a In r.Assignments
                If Not a Is Nothing And Not a.Task Is Nothing Then
                    Set t = a.Task
                    
                    ' Extraction des quantités depuis les champs personnalisés
                    On Error Resume Next
                    Select Case cfg.fieldPlan
                        Case "Number1": qtyPlan = a.Number1
                        Case "Number2": qtyPlan = a.Number2
                        Case "Number3": qtyPlan = a.Number3
                        Case "Number4": qtyPlan = a.Number4
                        Case "Number5": qtyPlan = a.Number5
                        Case Else: qtyPlan = 0
                    End Select
                    
                    Select Case cfg.fieldReal
                        Case "Number1": qtyReal = a.Number1
                        Case "Number2": qtyReal = a.Number2
                        Case "Number3": qtyReal = a.Number3
                        Case "Number4": qtyReal = a.Number4
                        Case "Number5": qtyReal = a.Number5
                        Case Else: qtyReal = 0
                    End Select
                    
                    Select Case cfg.fieldUnit
                        Case "Text1": unite = a.Text1
                        Case "Text2": unite = a.Text2
                        Case "Text3": unite = a.Text3
                        Case "Text4": unite = a.Text4
                        Case "Text5": unite = a.Text5
                        Case Else: unite = ""
                    End Select
                    On Error GoTo 0
                    
                    ' Calculs des écarts
                    ecartQty = qtyReal - qtyPlan
                    If qtyPlan > 0 Then
                        pctEcart = (ecartQty / qtyPlan) * 100
                    Else
                        pctEcart = 0
                    End If
                    
                    ' Calcul du drapeau
                    drapeauQty = ""
                    If Abs(pctEcart) >= cfg.seuilRelPct And Abs(ecartQty) >= cfg.seuilAbsQty Then
                        If ecartQty > 0 Then
                            drapeauQty = "Dépassement"
                        Else
                            drapeauQty = "Économie"
                        End If
                    End If
                    
                    ' Recherche du parent WBS
                    parentWbs = ""
                    If Not t.OutlineParent Is Nothing Then
                        parentWbs = t.OutlineParent.WBS
                    End If
                    
                    ' Construction du nom article
                    article = r.Name
                    If unite <> "" Then
                        article = article & " (" & unite & ")"
                    End If
                    
                    ' Export des données si des quantités existent
                    If qtyPlan > 0 Or qtyReal > 0 Then
                        row = row + 1
                        xlWs.Cells(row, 1) = article
                        xlWs.Cells(row, 2) = parentWbs
                        xlWs.Cells(row, 3) = t.Name
                        xlWs.Cells(row, 4) = qtyPlan
                        xlWs.Cells(row, 5) = qtyReal
                        xlWs.Cells(row, 6) = ecartQty
                        xlWs.Cells(row, 7) = pctEcart
                        xlWs.Cells(row, 8) = drapeauQty
                    End If
                End If
            Next a
        End If
    Next r
    
    ' Formatage
    xlWs.Range("D:F").NumberFormat = "0.0"
    xlWs.Range("G:G").NumberFormat = "0%"
    xlWs.Columns.AutoFit
    
    LogInfo "CONSOMMABLES créé avec " & (row - 1) & " lignes"
End Sub

' Export vers GUIDAGE
Private Sub WriteSheet_GUIDAGE(ByRef xlWb As Object, ByVal lastS0Path As String)
    Dim xlWs As Object
    Dim xlS0Wb As Object, xlS0Ws As Object, xlApp As Object
    Dim t As Task
    Dim row As Integer, s0Row As Integer, lastS0Row As Integer
    Dim existingData As Object
    Dim wbs As String, taskName As String, drapeaux As String
    
    Set xlWs = xlWb.Worksheets("GUIDAGE")
    Set existingData = CreateObject("Scripting.Dictionary")
    
    ' En-têtes
    row = 1
    xlWs.Cells(row, 1) = "WBS"
    xlWs.Cells(row, 2) = "ID"
    xlWs.Cells(row, 3) = "Tâche"
    xlWs.Cells(row, 4) = "Drapeaux"
    xlWs.Cells(row, 5) = "Levier"
    xlWs.Cells(row, 6) = "Responsable"
    xlWs.Cells(row, 7) = "Date_Action"
    xlWs.Cells(row, 8) = "Heures_récupérables_est"
    xlWs.Cells(row, 9) = "Statut"
    xlWs.Cells(row, 10) = "Commentaire"
    
    xlWs.Range("1:1").Font.Bold = True
    xlWs.Range("1:1").AutoFilter
    
    ' Chargement des données existantes depuis S0 si disponible
    If lastS0Path <> "" Then
        On Error Resume Next
        Set xlApp = CreateObject("Excel.Application")
        xlApp.Visible = False
        Set xlS0Wb = xlApp.Workbooks.Open(lastS0Path)
        Set xlS0Ws = xlS0Wb.Worksheets("GUIDAGE")
        
        lastS0Row = xlS0Ws.Cells(xlS0Ws.Rows.Count, 1).End(-4162).Row
        For s0Row = 2 To lastS0Row
            wbs = xlS0Ws.Cells(s0Row, 1).Value
            If wbs <> "" Then
                existingData(wbs) = Array( _
                    xlS0Ws.Cells(s0Row, 5).Value, _  ' Levier
                    xlS0Ws.Cells(s0Row, 6).Value, _  ' Responsable
                    xlS0Ws.Cells(s0Row, 7).Value, _  ' Date_Action
                    xlS0Ws.Cells(s0Row, 8).Value, _  ' Heures_récupérables_est
                    xlS0Ws.Cells(s0Row, 9).Value, _  ' Statut
                    xlS0Ws.Cells(s0Row, 10).Value _  ' Commentaire
                )
            End If
        Next s0Row
        
        xlS0Wb.Close False
        xlApp.Quit
        On Error GoTo 0
    End If
    
    ' Export des tâches avec possibilité de guidage
    For Each t In ActiveProject.Tasks
        If Not t Is Nothing And Not t.Summary Then
            wbs = t.WBS
            taskName = t.Name
            
            ' Détermination des drapeaux (logique simplifiée)
            drapeaux = ""
            If t.ActualWork > t.BaselineWork * 1.03 Then drapeaux = "Dérive"
            If t.ActualWork < t.BaselineWork * 0.97 Then drapeaux = "Économie"
            
            row = row + 1
            xlWs.Cells(row, 1) = wbs
            xlWs.Cells(row, 2) = t.ID
            xlWs.Cells(row, 3) = taskName
            xlWs.Cells(row, 4) = drapeaux
            
            ' Restauration des données existantes si disponibles
            If existingData.Exists(wbs) Then
                Dim savedData As Variant
                savedData = existingData(wbs)
                xlWs.Cells(row, 5) = savedData(0) ' Levier
                xlWs.Cells(row, 6) = savedData(1) ' Responsable
                xlWs.Cells(row, 7) = savedData(2) ' Date_Action
                xlWs.Cells(row, 8) = savedData(3) ' Heures_récupérables_est
                xlWs.Cells(row, 9) = savedData(4) ' Statut
                xlWs.Cells(row, 10) = savedData(5) ' Commentaire
            Else
                ' Valeurs par défaut pour nouvelles lignes
                xlWs.Cells(row, 5) = "" ' Levier
                xlWs.Cells(row, 6) = "" ' Responsable
                xlWs.Cells(row, 7) = "" ' Date_Action
                xlWs.Cells(row, 8) = 0 ' Heures_récupérables_est
                xlWs.Cells(row, 9) = "" ' Statut
                xlWs.Cells(row, 10) = "" ' Commentaire
            End If
        End If
    Next t
    
    ' Création des listes déroulantes
    Call CreateDropdowns(xlWs, row)
    
    ' Formatage
    xlWs.Range("H:H").NumberFormat = "0.0"
    xlWs.Columns.AutoFit
    
    LogInfo "GUIDAGE créé avec " & (row - 1) & " tâches"
End Sub

' Création des listes déroulantes pour GUIDAGE
Private Sub CreateDropdowns(ByRef xlWs As Object, ByVal lastRow As Integer)
    Dim levierList As String, statutList As String
    
    ' Listes pour les validations
    levierList = "Réallocation,Séquencement,Standardisation,Suppression doublons,Mutualisation déplacements,Autre"
    statutList = "À lancer,En cours,Clos"
    
    ' Application des validations (simplifiée)
    If lastRow > 1 Then
        ' Colonne Levier (E)
        With xlWs.Range("E2:E" & lastRow).Validation
            .Delete
            .Add Type:=3, AlertStyle:=1, Formula1:=levierList ' xlValidateList
        End With
        
        ' Colonne Statut (I)
        With xlWs.Range("I2:I" & lastRow).Validation
            .Delete
            .Add Type:=3, AlertStyle:=1, Formula1:=statutList ' xlValidateList
        End With
    End If
End Sub

' Export vers LOG
Private Sub WriteSheet_LOG(ByRef xlWb As Object)
    Dim xlWs As Object
    Dim row As Integer
    Dim logKey As Variant
    
    Set xlWs = xlWb.Worksheets("LOG")
    
    ' En-têtes
    row = 1
    xlWs.Cells(row, 1) = "Horodatage"
    xlWs.Cells(row, 2) = "Niveau"
    xlWs.Cells(row, 3) = "Message"
    
    xlWs.Range("1:1").Font.Bold = True
    xlWs.Range("1:1").AutoFilter
    
    ' Export des messages de log
    For Each logKey In logMessages.Keys
        row = row + 1
        xlWs.Cells(row, 1) = Format(Now, "dd/mm/yyyy hh:mm:ss")
        xlWs.Cells(row, 2) = Split(logMessages(logKey), "|")(0)
        xlWs.Cells(row, 3) = Split(logMessages(logKey), "|")(1)
    Next logKey
    
    ' Formatage
    xlWs.Columns.AutoFit
    
    LogInfo "LOG créé avec " & (row - 1) & " messages"
End Sub

' Fonctions de logging
Private Sub LogInfo(ByVal msg As String)
    logLevel = logLevel + 1
    logMessages("LOG" & logLevel) = "INFO|" & msg
End Sub

Private Sub LogWarn(ByVal msg As String)
    logLevel = logLevel + 1
    logMessages("LOG" & logLevel) = "WARN|" & msg
End Sub

Private Sub LogError(ByVal msg As String)
    logLevel = logLevel + 1
    logMessages("LOG" & logLevel) = "ERROR|" & msg
End Sub