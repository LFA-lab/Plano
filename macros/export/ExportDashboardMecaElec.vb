Option Explicit

' ============================================================================
'  CONFIGURATION : dossier de sortie des fichiers générés
'  Correspond au dossier du source VB sur la machine de dev (via WSL UNC)
' ============================================================================
Private Const OUTPUT_DIR As String = "\\wsl.localhost\Ubuntu\home\ntoi\LFA-lab\Plano\macros\export\"

' ============================================================================
'  MODULE : ExportDashboardMecaElec
'  OBJET  : Exporter un dashboard HTML style "Dashboard Interne" depuis MS Project
'           Style visuel identique à dashboard.html (bordures noires, barres jaunes)
'           Données filtrées par colonne H (Text4 = "Mecanique" ou "Electrique")
'  DATE   : 2025-02-03
' ============================================================================

' ============================================================================
'  ENTRY POINT : Fonction principale
' ============================================================================

Sub ExportDashboardMecaElec()
    On Error GoTo ErrorHandler
    
    Dim fso As Object, file As Object
    Dim filePath As String, logPath As String, dashboardTitle As String
    Dim t As Task
    
    ' Vérification du projet actif
    If ActiveProject Is Nothing Then
        MsgBox "Aucun projet actif !", vbCritical
        Exit Sub
    End If
    
    ' Récupération du nom du projet
    dashboardTitle = ActiveProject.Name
    For Each t In ActiveProject.Tasks
        If Not t Is Nothing Then
            dashboardTitle = t.Name
            Exit For
        End If
    Next t
    
    ' Chemins de sortie — dans le dossier du source VB (voir constante OUTPUT_DIR en haut)
    filePath = OUTPUT_DIR & "dashboard_mecaelec.html"
    logPath  = OUTPUT_DIR & "dashboard_mecaelec_DEBUG.txt"
    
    ' Génération du HTML complet
    Dim htmlContent As String
    htmlContent = BuildCompleteHTML(dashboardTitle)
    
    If Len(htmlContent) = 0 Then
        MsgBox "ERREUR : Le contenu HTML est vide !", vbCritical
        Exit Sub
    End If
    
    ' Génération du LOG de debug
    Dim logContent As String
    logContent = CreateDebugLog()
    
    ' Écriture des fichiers
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Fichier HTML
    Set file = fso.CreateTextFile(filePath, True, True)
    file.Write htmlContent
    file.Close
    
    ' Fichier LOG de debug
    Set file = fso.CreateTextFile(logPath, True, True)
    file.Write logContent
    file.Close
    
    MsgBox "Export termine !" & vbCrLf & vbCrLf & _
           "HTML : " & filePath & vbCrLf & _
           "LOG : " & logPath, vbInformation, "Export Dashboard"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "ERREUR #" & Err.Number & " :" & vbCrLf & Err.Description, vbCritical, "Erreur"
End Sub


' ============================================================================
'  CONSTRUCTION : HTML complet
' ============================================================================

Private Function BuildCompleteHTML(ByVal projectName As String) As String
    Dim html As String
    
    html = BuildHTMLHeader(projectName)
    html = html & BuildHTMLBody()
    html = html & BuildJavaScript()
    html = html & BuildHTMLFooter()
    
    BuildCompleteHTML = html
End Function


' ============================================================================
'  HEADER : CSS et structure de base
' ============================================================================

Private Function BuildHTMLHeader(ByVal projectName As String) As String
    Dim h As String
    
    h = "<!DOCTYPE html>" & vbCrLf
    h = h & "<html lang='fr'>" & vbCrLf
    h = h & "<head>" & vbCrLf
    h = h & "  <meta charset='UTF-8'>" & vbCrLf
    h = h & "  <meta name='viewport' content='width=device-width, initial-scale=1.0'>" & vbCrLf
    h = h & "  <title>Dashboard Interne - " & EncodeHTML(projectName) & "</title>" & vbCrLf
    h = h & BuildCSS()
    h = h & "</head>" & vbCrLf
    h = h & "<body>" & vbCrLf
    h = h & "<div class='dashboard-container'>" & vbCrLf
    h = h & "  <h1>Dashboard Interne - " & EncodeHTML(projectName) & "</h1>" & vbCrLf
    
    BuildHTMLHeader = h
End Function


' ============================================================================
'  CSS : Style identique à dashboard.html
' ============================================================================

Private Function BuildCSS() As String
    Dim css As String
    
    css = "<style>" & vbCrLf
    css = css & "* { margin: 0; padding: 0; box-sizing: border-box; }" & vbCrLf
    css = css & "body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background-color: #f5f5f5; padding: 20px; }" & vbCrLf
    css = css & ".dashboard-container { max-width: 1400px; margin: 0 auto; background: white; border: 2px solid #000; padding: 20px; }" & vbCrLf
    css = css & "h1 { font-size: 16px; font-weight: bold; margin-bottom: 20px; text-align: left; }" & vbCrLf
    
    ' Onglets
    css = css & ".tabs-row { display: flex; gap: 8px; flex-wrap: wrap; margin-bottom: 15px; }" & vbCrLf
    css = css & ".tab-button { padding: 8px 16px; background: white; border: 1px solid #000; cursor: pointer; font-size: 11px; font-weight: 600; transition: all 0.2s; }" & vbCrLf
    css = css & ".tab-button:hover { background: #e8e8e8; }" & vbCrLf
    css = css & ".tab-button.active { background: #d0d0d0; border-width: 2px; }" & vbCrLf
    
    ' Sections de vue
    css = css & ".view-section { display: none; }" & vbCrLf
    css = css & ".view-section.active { display: block; }" & vbCrLf
    
    ' Vue mécanique - Barres de progression
    css = css & ".mechanical-progress-view { padding: 20px; border: 1px solid #000; }" & vbCrLf
    css = css & ".mechanical-progress-title { text-align: center; font-size: 16px; font-weight: bold; margin-bottom: 30px; }" & vbCrLf
    css = css & ".mechanical-progress-grid { max-width: 900px; margin: 0 auto; }" & vbCrLf
    css = css & ".mechanical-progress-row { display: grid; grid-template-columns: 180px 1fr 80px; align-items: center; gap: 10px; margin-bottom: 8px; }" & vbCrLf
    css = css & ".mechanical-progress-label { font-size: 9px; font-weight: 600; text-align: left; line-height: 1.2; }" & vbCrLf
    css = css & ".mechanical-progress-bar-container { background: white; border: 2px solid #000; height: 16px; position: relative; }" & vbCrLf
    css = css & ".mechanical-progress-bar { height: 100%; background: #f4b400; border-right: 1px solid #000; transition: width 0.3s ease; }" & vbCrLf
    css = css & ".mechanical-progress-percentage { font-size: 9px; font-weight: bold; text-align: right; }" & vbCrLf
    css = css & ".mechanical-progress-row.general { margin-top: 15px; padding-top: 12px; border-top: 2px solid #000; }" & vbCrLf
    css = css & ".mechanical-progress-row.general .mechanical-progress-label { font-size: 10px; font-weight: bold; }" & vbCrLf
    css = css & ".mechanical-progress-row.general .mechanical-progress-percentage { font-size: 10px; }" & vbCrLf
    
    ' Vue électrique - Tableau
    css = css & ".electrical-view { padding: 20px; }" & vbCrLf
    css = css & ".progress-table { width: 100%; border-collapse: collapse; border: 1px solid #000; margin-bottom: 20px; }" & vbCrLf
    css = css & ".progress-table th, .progress-table td { border: 1px solid #000; padding: 8px; text-align: center; font-size: 11px; }" & vbCrLf
    css = css & ".progress-table th { background: #e0e0e0; font-weight: 600; }" & vbCrLf
    css = css & ".progress-table td.action-name { text-align: left; font-weight: 600; background: #f5f5f5; }" & vbCrLf
    css = css & ".progress-table td.percentage { font-weight: 600; }" & vbCrLf
    css = css & ".progress-table td.percentage.low { color: #d32f2f; }" & vbCrLf
    css = css & ".progress-table td.percentage.medium { color: #f57c00; }" & vbCrLf
    css = css & ".progress-table td.percentage.high { color: #388e3c; }" & vbCrLf
    
    css = css & "</style>" & vbCrLf
    
    BuildCSS = css
End Function


' ============================================================================
'  BODY : Contenu principal avec onglets et vues
' ============================================================================

Private Function BuildHTMLBody() As String
    Dim html As String
    
    ' Onglets de navigation
    html = "<div class='tabs-row'>" & vbCrLf
    html = html & "  <button class='tab-button active' data-view='metiers'>Vue Métiers</button>" & vbCrLf
    html = html & "  <button class='tab-button' data-view='client'>Dashboard Client</button>" & vbCrLf
    html = html & "</div>" & vbCrLf
    
    ' Vue Métiers — le sélecteur de sous-zone est encapsulé DANS cette vue
    html = html & BuildViewMetiers()
    
    ' Vue Client
    html = html & BuildViewClient()
    
    BuildHTMLBody = html
End Function


' ============================================================================
'  SÉLECTEUR DE ZONE
' ============================================================================

Private Function BuildZoneSelector(Optional ByVal selectId As String = "zoneFilter", _
                                    Optional ByVal infoId As String = "zoneInfo") As String
    Dim html As String
    Dim t As Task
    Dim zone As String
    Dim zones As Object
    Dim zoneKey As Variant
    Dim sortedZones() As String
    Dim i As Long, j As Long, temp As String
    
    Set zones = CreateObject("Scripting.Dictionary")
    
    ' Collecter toutes les sous-zones uniques depuis les tâches feuilles (Text3)
    For Each t In ActiveProject.Tasks
        If Not t Is Nothing And Not t.Summary Then
            zone = IIf(Len(t.Text3) > 0, CStr(t.Text3), "")
            If Len(zone) > 0 And Not zones.Exists(zone) Then
                zones.Add zone, True
            End If
        End If
    Next t
    
    ' Trier les zones par ordre alphabétique
    If zones.Count > 0 Then
        ReDim sortedZones(zones.Count - 1)
        i = 0
        For Each zoneKey In zones.Keys
            sortedZones(i) = CStr(zoneKey)
            i = i + 1
        Next zoneKey
        
        For i = 0 To UBound(sortedZones) - 1
            For j = i + 1 To UBound(sortedZones)
                If sortedZones(i) > sortedZones(j) Then
                    temp = sortedZones(i)
                    sortedZones(i) = sortedZones(j)
                    sortedZones(j) = temp
                End If
            Next j
        Next i
    End If
    
    html = "<div style='margin:15px 0;padding:12px;border:2px solid #000;background:#f9f9f9;'>" & vbCrLf
    html = html & "  <label style='font-size:11px;font-weight:600;margin-right:10px;'>Affichage par sous-zone :</label>" & vbCrLf
    html = html & "  <select id='" & selectId & "' style='padding:6px 12px;font-size:11px;border:2px solid #000;background:white;cursor:pointer;min-width:200px;'>" & vbCrLf
    html = html & "    <option value='all' selected>&#10003; Toutes les zones (consolidé)</option>" & vbCrLf
    
    If zones.Count > 0 Then
        For i = 0 To UBound(sortedZones)
            html = html & "    <option value='" & EncodeHTML(sortedZones(i)) & "'>" & EncodeHTML(sortedZones(i)) & "</option>" & vbCrLf
        Next i
    End If
    
    html = html & "  </select>" & vbCrLf
    html = html & "  <span id='" & infoId & "' style='margin-left:15px;font-size:10px;color:#666;'></span>" & vbCrLf
    html = html & "</div>" & vbCrLf
    
    BuildZoneSelector = html
End Function


' ============================================================================
'  VUE MÉTIERS : VRD/MECA/ELEC en 3 colonnes
' ============================================================================

Private Function BuildViewMetiers() As String
    Dim html As String
    
    html = "<div class='view-section active' data-view='metiers'>" & vbCrLf
    ' Sélecteur de sous-zone : encapsulé dans la vue Métiers uniquement
    html = html & BuildZoneSelector()
    html = html & "  <div style='display:grid;grid-template-columns:repeat(3,1fr);gap:15px;padding:15px;'>" & vbCrLf
    
    ' Section TRANCHÉES (VRD)
    html = html & "    <div style='border:2px solid #000;padding:12px;background:#fff;'>" & vbCrLf
    html = html & "      <h3 style='text-align:center;font-weight:bold;margin-bottom:15px;font-size:13px;'>TRANCHÉES</h3>" & vbCrLf
    html = html & BuildMetierSection("vrd")
    html = html & "    </div>" & vbCrLf
    
    ' Section STRUCTURES (MECA)
    html = html & "    <div style='border:2px solid #000;padding:12px;background:#fff;'>" & vbCrLf
    html = html & "      <h3 style='text-align:center;font-weight:bold;margin-bottom:15px;font-size:13px;'>STRUCTURES</h3>" & vbCrLf
    html = html & BuildMetierSection("meca")
    html = html & "    </div>" & vbCrLf
    
    ' Section ÉLECTRICITÉ (ELEC)
    html = html & "    <div style='border:2px solid #000;padding:12px;background:#fff;'>" & vbCrLf
    html = html & "      <h3 style='text-align:center;font-weight:bold;margin-bottom:15px;font-size:13px;'>ÉLECTRICITÉ</h3>" & vbCrLf
    html = html & BuildMetierSection("elec")
    html = html & "    </div>" & vbCrLf
    
    html = html & "  </div>" & vbCrLf
    html = html & "</div>" & vbCrLf
    
    BuildViewMetiers = html
End Function


' ============================================================================
'  HELPER : Générer une section métier (VRD/MECA/ELEC)
'
'  Approche : itérer directement sur les tâches FEUILLES filtrées par Text4.
'  Chaque tâche feuille devient une ligne indépendante (son nom propre, sa zone,
'  son % MSProject direct). Pas de groupement par récapitulatif.
'  "Avancement général" = moyenne pondérée par durée de toutes les feuilles du métier.
' ============================================================================

Private Function BuildMetierSection(ByVal metierType As String) As String
    Dim html         As String
    Dim t            As Task
    Dim lot          As String
    Dim subZone      As String
    Dim dur          As Double
    Dim pct          As Double
    Dim avgPct       As Double
    Dim pctStr       As String
    Dim genSumPctDur As Double
    Dim genSumDur    As Double
    Dim genCount     As Long
    Dim genSumPct    As Double
    Dim hasRows      As Boolean
    
    hasRows = False
    
    ' ----------------------------------------------------------------
    ' Itération unique sur toutes les tâches feuilles
    ' ----------------------------------------------------------------
    For Each t In ActiveProject.Tasks
        If Not t Is Nothing And Not t.Summary Then
            lot = LCase$(Trim$(IIf(Len(t.Text4) > 0, CStr(t.Text4), "")))
            
            If MetierMatch(metierType, lot) Then
                subZone = IIf(Len(t.Text3) > 0, CStr(t.Text3), "SANS_ZONE")
                
                dur = 0
                On Error Resume Next
                If Not IsEmpty(t.Duration) And Not IsNull(t.Duration) Then
                    dur = CDbl(t.Duration)
                End If
                On Error GoTo 0
                
                pct = 0
                On Error Resume Next
                If Not IsEmpty(t.PercentComplete) And Not IsNull(t.PercentComplete) Then
                    pct = CDbl(t.PercentComplete)
                End If
                On Error GoTo 0
                
                ' Accumuler pour l'avancement général
                genSumPctDur = genSumPctDur + pct * dur
                genSumDur    = genSumDur    + dur
                genCount     = genCount     + 1
                genSumPct    = genSumPct    + pct
                
                ' Ligne HTML : nom réel de la tâche feuille
                pctStr = Replace(CStr(pct), ",", ".")
                html = html & "      <div class='mechanical-progress-row' data-zone='" & EncodeHTML(subZone) & "' data-pct='" & pctStr & "'>" & vbCrLf
                html = html & "        <div class='mechanical-progress-label'>" & EncodeHTML(t.Name) & "</div>" & vbCrLf
                html = html & "        <div class='mechanical-progress-bar-container'>" & vbCrLf
                html = html & "          <div class='mechanical-progress-bar' style='width:" & pctStr & "%;'></div>" & vbCrLf
                html = html & "        </div>" & vbCrLf
                html = html & "        <div class='mechanical-progress-percentage'>" & Replace(CStr(Round(pct, 1)), ".", ",") & "%</div>" & vbCrLf
                html = html & "      </div>" & vbCrLf
                
                hasRows = True
            End If
        End If
    Next t
    
    ' ----------------------------------------------------------------
    ' Ligne "Avancement général" — pondérée par durée, fallback simple
    ' ----------------------------------------------------------------
    If hasRows Then
        If genSumDur > 0 Then
            avgPct = genSumPctDur / genSumDur
        ElseIf genCount > 0 Then
            avgPct = genSumPct / genCount
        Else
            avgPct = 0
        End If
        
        pctStr = Replace(CStr(avgPct), ",", ".")
        html = html & "      <div class='mechanical-progress-row general' data-zone='__GENERAL__' data-pct='" & pctStr & "'>" & vbCrLf
        html = html & "        <div class='mechanical-progress-label'>Avancement général</div>" & vbCrLf
        html = html & "        <div class='mechanical-progress-bar-container'>" & vbCrLf
        html = html & "          <div class='mechanical-progress-bar' style='width:" & pctStr & "%;'></div>" & vbCrLf
        html = html & "        </div>" & vbCrLf
        html = html & "        <div class='mechanical-progress-percentage'>" & Replace(CStr(Round(avgPct, 1)), ".", ",") & "%</div>" & vbCrLf
        html = html & "      </div>" & vbCrLf
    Else
        html = html & "      <p style='text-align:center;color:#999;font-size:9px;padding:12px;'>Aucune donnée</p>" & vbCrLf
    End If
    
    BuildMetierSection = html
End Function


' ============================================================================
'  HELPER : Vérifier si une tâche est sous-tâche d'une autre
' ============================================================================

Private Function IsSubTaskOf(ByVal childTask As Task, ByVal parentTask As Task) As Boolean
    Dim parentOutline As String
    Dim childOutline As String
    
    IsSubTaskOf = False
    If childTask Is Nothing Or parentTask Is Nothing Then Exit Function
    
    ' Vérifier le niveau de plan
    If childTask.OutlineLevel <= parentTask.OutlineLevel Then Exit Function
    
    ' Vérifier le numéro de plan (commence par le numéro du parent)
    parentOutline = parentTask.OutlineNumber
    childOutline = childTask.OutlineNumber
    
    If Len(childOutline) > Len(parentOutline) Then
        If Left(childOutline, Len(parentOutline)) = parentOutline Then
            IsSubTaskOf = True
        End If
    End If
End Function


' ============================================================================
'  HELPER : Tester si un lot (Text4 en minuscules/trimé) correspond au métier
' ============================================================================

Private Function MetierMatch(ByVal metierType As String, ByVal lotLower As String) As Boolean
    Select Case metierType
        Case "vrd"
            MetierMatch = (InStr(lotLower, "vrd") > 0)
        Case "meca"
            MetierMatch = (InStr(lotLower, "meca") > 0 Or InStr(lotLower, "mecanique") > 0)
        Case "elec"
            MetierMatch = (InStr(lotLower, "elec") > 0 Or InStr(lotLower, "electrique") > 0)
        Case Else
            MetierMatch = False
    End Select
End Function


' ============================================================================
'  JAVASCRIPT : Gestion des onglets
' ============================================================================

Private Function BuildJavaScript() As String
    Dim js As String
    
    js = "<script>" & vbCrLf
    js = js & "document.addEventListener('DOMContentLoaded', function() {" & vbCrLf
    js = js & "  const buttons = document.querySelectorAll('.tab-button');" & vbCrLf
    js = js & "  const views = document.querySelectorAll('.view-section');" & vbCrLf
    js = js & "  const zoneFilter = document.getElementById('zoneFilter');" & vbCrLf
    js = js & "  const zoneInfo = document.getElementById('zoneInfo');" & vbCrLf
    js = js & vbCrLf
    js = js & "  // Gestion des onglets" & vbCrLf
    js = js & "  buttons.forEach(btn => {" & vbCrLf
    js = js & "    btn.addEventListener('click', function() {" & vbCrLf
    js = js & "      const targetView = this.getAttribute('data-view');" & vbCrLf
    js = js & "      buttons.forEach(b => b.classList.remove('active'));" & vbCrLf
    js = js & "      this.classList.add('active');" & vbCrLf
    js = js & "      views.forEach(v => {" & vbCrLf
    js = js & "        if (v.getAttribute('data-view') === targetView) {" & vbCrLf
    js = js & "          v.classList.add('active');" & vbCrLf
    js = js & "        } else {" & vbCrLf
    js = js & "          v.classList.remove('active');" & vbCrLf
    js = js & "        }" & vbCrLf
    js = js & "      });" & vbCrLf
    js = js & "    });" & vbCrLf
    js = js & "  });" & vbCrLf
    js = js & vbCrLf
    js = js & "  // Gestion du filtrage par zone — Vue Métiers" & vbCrLf
    js = js & "  if (zoneFilter) {" & vbCrLf
    js = js & "    zoneFilter.addEventListener('change', function() {" & vbCrLf
    js = js & "      const selectedZone = this.value;" & vbCrLf
    js = js & "      const rows = document.querySelectorAll('[data-view=""metiers""] [data-zone]');" & vbCrLf
    js = js & "      let visibleCount = 0;" & vbCrLf
    js = js & "      let totalCount = 0;" & vbCrLf
    js = js & "      let sumPct = 0;" & vbCrLf
    js = js & "      let visibleDataRows = 0;" & vbCrLf
    js = js & vbCrLf
    js = js & "      rows.forEach(row => {" & vbCrLf
    js = js & "        const rowZone = row.getAttribute('data-zone');" & vbCrLf
    js = js & "        const rowPct = parseFloat(row.getAttribute('data-pct')) || 0;" & vbCrLf
    js = js & "        " & vbCrLf
    js = js & "        if (rowZone !== '__GENERAL__') totalCount++;" & vbCrLf
    js = js & "        " & vbCrLf
    js = js & "        if (selectedZone === 'all') {" & vbCrLf
    js = js & "          row.style.display = '';" & vbCrLf
    js = js & "          visibleCount++;" & vbCrLf
    js = js & "          if (rowZone !== '__GENERAL__') {" & vbCrLf
    js = js & "            sumPct += rowPct;" & vbCrLf
    js = js & "            visibleDataRows++;" & vbCrLf
    js = js & "          }" & vbCrLf
    js = js & "        } else {" & vbCrLf
    js = js & "          if (rowZone === selectedZone) {" & vbCrLf
    js = js & "            row.style.display = '';" & vbCrLf
    js = js & "            visibleCount++;" & vbCrLf
    js = js & "            sumPct += rowPct;" & vbCrLf
    js = js & "            visibleDataRows++;" & vbCrLf
    js = js & "          } else if (rowZone === '__GENERAL__') {" & vbCrLf
    js = js & "            row.style.display = 'none';" & vbCrLf
    js = js & "          } else {" & vbCrLf
    js = js & "            row.style.display = 'none';" & vbCrLf
    js = js & "          }" & vbCrLf
    js = js & "        }" & vbCrLf
    js = js & "      });" & vbCrLf
    js = js & vbCrLf
    js = js & "      // Recalculer avancement général" & vbCrLf
    js = js & "      const generalRows = document.querySelectorAll('[data-view=""metiers""] .mechanical-progress-row.general');" & vbCrLf
    js = js & "      generalRows.forEach(generalRow => {" & vbCrLf
    js = js & "        if (selectedZone !== 'all') {" & vbCrLf
    js = js & "          const avgPct = visibleDataRows > 0 ? sumPct / visibleDataRows : 0;" & vbCrLf
    js = js & "          const bar = generalRow.querySelector('.mechanical-progress-bar');" & vbCrLf
    js = js & "          const percentage = generalRow.querySelector('.mechanical-progress-percentage');" & vbCrLf
    js = js & "          const label = generalRow.querySelector('.mechanical-progress-label');" & vbCrLf
    js = js & "          if (bar) bar.style.width = avgPct + '%';" & vbCrLf
    js = js & "          if (percentage) percentage.textContent = avgPct.toFixed(1).replace('.', ',') + '%';" & vbCrLf
    js = js & "          if (label) label.textContent = 'Avancement ' + selectedZone;" & vbCrLf
    js = js & "          generalRow.style.display = '';" & vbCrLf
    js = js & "          generalRow.setAttribute('data-zone', selectedZone);" & vbCrLf
    js = js & "        } else {" & vbCrLf
    js = js & "          // Retour à 'Toutes les zones' : restaurer l'état initial de la ligne générale" & vbCrLf
    js = js & "          const origPct = parseFloat(generalRow.getAttribute('data-pct')) || 0;" & vbCrLf
    js = js & "          generalRow.setAttribute('data-zone', '__GENERAL__');" & vbCrLf
    js = js & "          generalRow.style.display = '';" & vbCrLf
    js = js & "          const bar = generalRow.querySelector('.mechanical-progress-bar');" & vbCrLf
    js = js & "          const percentage = generalRow.querySelector('.mechanical-progress-percentage');" & vbCrLf
    js = js & "          const label = generalRow.querySelector('.mechanical-progress-label');" & vbCrLf
    js = js & "          if (bar) bar.style.width = origPct + '%';" & vbCrLf
    js = js & "          if (percentage) percentage.textContent = origPct.toFixed(1).replace('.', ',') + '%';" & vbCrLf
    js = js & "          if (label) label.textContent = 'Avancement général';" & vbCrLf
    js = js & "        }" & vbCrLf
    js = js & "      });" & vbCrLf
    js = js & vbCrLf
    js = js & "      // Mise à jour des informations" & vbCrLf
    js = js & "      if (zoneInfo) {" & vbCrLf
    js = js & "        if (selectedZone === 'all') {" & vbCrLf
    js = js & "          zoneInfo.textContent = totalCount + ' ligne' + (totalCount > 1 ? 's' : '') + ' (toutes zones agrégées)';" & vbCrLf
    js = js & "        } else {" & vbCrLf
    js = js & "          zoneInfo.textContent = visibleDataRows + ' ligne' + (visibleDataRows > 1 ? 's' : '') + ' affichée' + (visibleDataRows > 1 ? 's' : '');" & vbCrLf
    js = js & "        }" & vbCrLf
    js = js & "      }" & vbCrLf
    js = js & "    });" & vbCrLf
    js = js & vbCrLf
    js = js & "    // Initialiser le compteur" & vbCrLf
    js = js & "    const totalRows = document.querySelectorAll('[data-view=""metiers""] [data-zone]:not([data-zone=""__GENERAL__""])').length;" & vbCrLf
    js = js & "    if (zoneInfo) {" & vbCrLf
    js = js & "      zoneInfo.textContent = totalRows + ' ligne' + (totalRows > 1 ? 's' : '') + ' (toutes zones agrégées)';" & vbCrLf
    js = js & "    }" & vbCrLf
    js = js & "  }" & vbCrLf
    js = js & vbCrLf
    js = js & "  // Gestion du filtrage par zone — Dashboard Client" & vbCrLf
    js = js & "  const zoneFilterClient = document.getElementById('zoneFilterClient');" & vbCrLf
    js = js & "  const zoneInfoClient   = document.getElementById('zoneInfoClient');" & vbCrLf
    js = js & "  if (zoneFilterClient) {" & vbCrLf
    js = js & "    zoneFilterClient.addEventListener('change', function() {" & vbCrLf
    js = js & "      const selectedZone = this.value;" & vbCrLf
    js = js & "      const rows = document.querySelectorAll('[data-view=""client""] [data-zone]');" & vbCrLf
    js = js & "      let visible = 0;" & vbCrLf
    js = js & "      rows.forEach(row => {" & vbCrLf
    js = js & "        const rz = row.getAttribute('data-zone');" & vbCrLf
    js = js & "        if (selectedZone === 'all' || rz === selectedZone) {" & vbCrLf
    js = js & "          row.style.display = '';" & vbCrLf
    js = js & "          visible++;" & vbCrLf
    js = js & "        } else {" & vbCrLf
    js = js & "          row.style.display = 'none';" & vbCrLf
    js = js & "        }" & vbCrLf
    js = js & "      });" & vbCrLf
    js = js & "      if (zoneInfoClient) {" & vbCrLf
    js = js & "        if (selectedZone === 'all') {" & vbCrLf
    js = js & "          zoneInfoClient.textContent = rows.length + ' lot' + (rows.length > 1 ? 's' : '') + ' (toutes zones)';" & vbCrLf
    js = js & "        } else {" & vbCrLf
    js = js & "          zoneInfoClient.textContent = visible + ' lot' + (visible > 1 ? 's' : '') + ' affiché' + (visible > 1 ? 's' : '');" & vbCrLf
    js = js & "        }" & vbCrLf
    js = js & "      }" & vbCrLf
    js = js & "    });" & vbCrLf
    js = js & "    // Init compteur Dashboard Client" & vbCrLf
    js = js & "    const totalClient = document.querySelectorAll('[data-view=""client""] [data-zone]').length;" & vbCrLf
    js = js & "    if (zoneInfoClient) {" & vbCrLf
    js = js & "      zoneInfoClient.textContent = totalClient + ' lot' + (totalClient > 1 ? 's' : '') + ' (toutes zones)';" & vbCrLf
    js = js & "    }" & vbCrLf
    js = js & "  }" & vbCrLf
    js = js & "});" & vbCrLf
    js = js & "</script>" & vbCrLf
    
    BuildJavaScript = js
End Function


' ============================================================================
'  VUE CLIENT : Dashboard Client EDF RE
' ============================================================================

Private Function BuildViewClient() As String
    Dim html As String
    
    html = "<div class='view-section' data-view='client'>" & vbCrLf
    html = html & "  <h2 style='text-align:center;margin-bottom:20px;'>DASHBOARD CLIENT - EDF RE</h2>" & vbCrLf
    
    ' Sélecteur de zone — même logique que Vue Métiers, IDs distincts
    html = html & BuildZoneSelector("zoneFilterClient", "zoneInfoClient")
    
    ' Section 1 : Histogramme par grande catégorie
    html = html & BuildClientSection1_Histogramme()
    html = html & "<div style='margin:40px 0;border-top:2px solid #ddd;'></div>" & vbCrLf
    
    ' Section 2 : 3 Prochaines tâches à finir
    html = html & BuildClientSection2_ProchainesFins()
    html = html & "<div style='margin:40px 0;border-top:2px solid #ddd;'></div>" & vbCrLf
    
    ' Section 3 : 3 Prochains démarrages
    html = html & BuildClientSection3_ProchainsStarts()
    html = html & "<div style='margin:40px 0;border-top:2px solid #ddd;'></div>" & vbCrLf
    
    ' Section 4 : Placeholder avancement semaine
    html = html & BuildClientSection4_AvancementSemaine()
    
    html = html & "</div>" & vbCrLf
    
    BuildViewClient = html
End Function


' ============================================================================
'  CLIENT SECTION 1 : Histogramme % avancement par grande catégorie
' ============================================================================

Private Function BuildClientSection1_Histogramme() As String
    Dim html     As String
    Dim t        As Task
    Dim child    As Task
    Dim summaryTasks As Object
    Dim pct      As Double
    Dim i As Long, j As Long
    Dim k As Variant, temp As Variant
    Dim taskName As String, taskPct As Double, taskZone As String
    Dim recapZone As String
    Dim pctStr   As String
    
    Set summaryTasks = CreateObject("Scripting.Dictionary")
    
    html = "<h3 style='font-size:14px;font-weight:bold;margin-bottom:20px;'>AVANCEMENT PAR GRANDE CATÉGORIE</h3>" & vbCrLf
    html = html & "<div class='mechanical-progress-grid' style='max-width:900px;margin:0 auto;'>" & vbCrLf
    
    ' Récapitulatifs de niveau 2 uniquement (lots directs : VRD, Meca, MALT…)
    For Each t In ActiveProject.Tasks
        On Error Resume Next
        If Not t Is Nothing Then
            If t.Summary And t.OutlineLevel = 2 Then
                pct = 0
                If Not IsEmpty(t.PercentComplete) And Not IsNull(t.PercentComplete) Then
                    pct = CDbl(t.PercentComplete)
                End If
                
                ' Détecter la zone : première tâche feuille enfant
                recapZone = "SANS_ZONE"
                For Each child In ActiveProject.Tasks
                    If Not child Is Nothing And Not child.Summary Then
                        If IsSubTaskOf(child, t) Then
                            recapZone = IIf(Len(child.Text3) > 0, CStr(child.Text3), "SANS_ZONE")
                            Exit For
                        End If
                    End If
                Next child
                
                summaryTasks.Add t.ID, Array(t.Name, pct, recapZone)
            End If
        End If
        On Error GoTo 0
    Next t
    
    ' Trier par % décroissant
    If summaryTasks.Count > 0 Then
        Dim keys() As Variant
        ReDim keys(summaryTasks.Count - 1)
        i = 0
        For Each k In summaryTasks.Keys
            keys(i) = k
            i = i + 1
        Next k
        
        For i = 0 To UBound(keys) - 1
            For j = i + 1 To UBound(keys)
                If summaryTasks(keys(i))(1) < summaryTasks(keys(j))(1) Then
                    temp = keys(i)
                    keys(i) = keys(j)
                    keys(j) = temp
                End If
            Next j
        Next i
        
        ' Générer les barres avec data-zone + data-pct pour le filtre JS
        For i = 0 To UBound(keys)
            taskName = summaryTasks(keys(i))(0)
            taskPct  = summaryTasks(keys(i))(1)
            taskZone = CStr(summaryTasks(keys(i))(2))
            pctStr   = Replace(CStr(taskPct), ",", ".")
            
            html = html & "  <div class='mechanical-progress-row' data-zone='" & EncodeHTML(taskZone) & "' data-pct='" & pctStr & "'>" & vbCrLf
            html = html & "    <div class='mechanical-progress-label'>" & EncodeHTML(taskName) & "</div>" & vbCrLf
            html = html & "    <div class='mechanical-progress-bar-container'>" & vbCrLf
            html = html & "      <div class='mechanical-progress-bar' style='width:" & pctStr & "%;'></div>" & vbCrLf
            html = html & "    </div>" & vbCrLf
            html = html & "    <div class='mechanical-progress-percentage'>" & Replace(CStr(Round(taskPct, 1)), ".", ",") & "%</div>" & vbCrLf
            html = html & "  </div>" & vbCrLf
        Next i
    Else
        html = html & "  <p style='text-align:center;color:#666;'>Aucune tâche récapitulative trouvée</p>" & vbCrLf
    End If
    
    html = html & "</div>" & vbCrLf
    
    BuildClientSection1_Histogramme = html
End Function


' ============================================================================
'  CLIENT SECTION 2 : 3 Prochaines tâches à finir
' ============================================================================

Private Function BuildClientSection2_ProchainesFins() As String
    Dim html As String
    Dim t As Task
    Dim candidateTasks As Object
    Dim dateToday As Date
    Dim pct As Double
    Dim sortKey As String
    Dim i As Long
    Dim k As Variant
    Dim j As Long, temp As Variant
    Dim maxRows As Long
    Dim taskName As String, taskPct As Double, taskFinish As Date
    Dim pctClass As String
    
    Set candidateTasks = CreateObject("Scripting.Dictionary")
    dateToday = Date
    
    html = "<h3 style='font-size:14px;font-weight:bold;margin:40px 0 20px 0;'>PROCHAINES TÂCHES À TERMINER</h3>" & vbCrLf
    
    ' Collecter les tâches en cours (0 < % < 100, Finish >= Aujourd'hui)
    For Each t In ActiveProject.Tasks
        On Error Resume Next
        If Not t Is Nothing Then
            If Not t.Summary Then
                pct = 0
                
                If Not IsEmpty(t.PercentComplete) And Not IsNull(t.PercentComplete) Then
                    pct = CDbl(t.PercentComplete)
                End If
                
                If pct > 0 And pct < 100 Then
                    If Not IsEmpty(t.Finish) And Not IsNull(t.Finish) Then
                        If t.Finish >= dateToday Then
                            ' Clé de tri : Date_Finish (YYYYMMDD) + % inversé (pour tri décroissant si même date)
                            sortKey = Format(t.Finish, "yyyymmdd") & "_" & Format(1000 - pct, "0000")
                            candidateTasks.Add sortKey, Array(t.Name, pct, t.Finish)
                        End If
                    End If
                End If
            End If
        End If
        On Error GoTo 0
    Next t
    
    ' Trier et limiter à 3
    html = html & "<table class='progress-table'>" & vbCrLf
    html = html & "  <thead>" & vbCrLf
    html = html & "    <tr>" & vbCrLf
    html = html & "      <th>Tâche</th>" & vbCrLf
    html = html & "      <th>% Avancement</th>" & vbCrLf
    html = html & "      <th>Date fin prévue</th>" & vbCrLf
    html = html & "    </tr>" & vbCrLf
    html = html & "  </thead>" & vbCrLf
    html = html & "  <tbody>" & vbCrLf
    
    If candidateTasks.Count > 0 Then
        ' Trier les clés (ordre alphabétique = ordre chronologique)
        Dim keys() As Variant
        ReDim keys(candidateTasks.Count - 1)
        
        i = 0
        For Each k In candidateTasks.Keys
            keys(i) = k
            i = i + 1
        Next k
        
        ' Tri à bulles
        For i = 0 To UBound(keys) - 1
            For j = i + 1 To UBound(keys)
                If keys(i) > keys(j) Then
                    temp = keys(i)
                    keys(i) = keys(j)
                    keys(j) = temp
                End If
            Next j
        Next i
        
        ' Afficher les 3 premières
        maxRows = 3
        If UBound(keys) + 1 < maxRows Then maxRows = UBound(keys) + 1
        
        For i = 0 To maxRows - 1
            taskName = candidateTasks(keys(i))(0)
            taskPct = candidateTasks(keys(i))(1)
            taskFinish = candidateTasks(keys(i))(2)
            
            If taskPct < 50 Then
                pctClass = "low"
            ElseIf taskPct < 80 Then
                pctClass = "medium"
            Else
                pctClass = "high"
            End If
            
            html = html & "    <tr>" & vbCrLf
            html = html & "      <td class='action-name'>" & EncodeHTML(taskName) & "</td>" & vbCrLf
            html = html & "      <td class='percentage " & pctClass & "'>" & Replace(CStr(Round(taskPct, 1)), ".", ",") & "%</td>" & vbCrLf
            html = html & "      <td>" & Format(taskFinish, "dd/mm/yyyy") & "</td>" & vbCrLf
            html = html & "    </tr>" & vbCrLf
        Next i
    Else
        html = html & "    <tr><td colspan='3' style='text-align:center;color:#666;'>Aucune tâche en cours à terminer</td></tr>" & vbCrLf
    End If
    
    html = html & "  </tbody>" & vbCrLf
    html = html & "</table>" & vbCrLf
    
    BuildClientSection2_ProchainesFins = html
End Function


' ============================================================================
'  CLIENT SECTION 3 : 3 Prochaines tâches à démarrer
' ============================================================================

Private Function BuildClientSection3_ProchainsStarts() As String
    Dim html As String
    Dim t As Task
    Dim candidateTasks As Object
    Dim dateToday As Date
    Dim pct As Double
    Dim sortKey As String
    Dim uniqueKey As String
    Dim counter As Integer
    Dim i As Long
    Dim k As Variant
    Dim j As Long, temp As Variant
    Dim maxRows As Long
    Dim taskName As String, taskStart As Date
    
    Set candidateTasks = CreateObject("Scripting.Dictionary")
    dateToday = Date
    
    html = "<h3 style='font-size:14px;font-weight:bold;margin:40px 0 20px 0;'>PROCHAINES TÂCHES À DÉMARRER</h3>" & vbCrLf
    
    ' Collecter les tâches non démarrées (% = 0, Start >= Aujourd'hui)
    For Each t In ActiveProject.Tasks
        On Error Resume Next
        If Not t Is Nothing Then
            If Not t.Summary Then
                pct = 0
                
                If Not IsEmpty(t.PercentComplete) And Not IsNull(t.PercentComplete) Then
                    pct = CDbl(t.PercentComplete)
                End If
                
                If pct = 0 Then
                    If Not IsEmpty(t.Start) And Not IsNull(t.Start) Then
                        If t.Start >= dateToday Then
                            ' Clé de tri : Date_Start (YYYYMMDD)
                            sortKey = Format(t.Start, "yyyymmdd")
                            
                            ' Gérer les doublons de clé
                            uniqueKey = sortKey
                            counter = 0
                            Do While candidateTasks.Exists(uniqueKey)
                                counter = counter + 1
                                uniqueKey = sortKey & "_" & Format(counter, "0000")
                            Loop
                            
                            candidateTasks.Add uniqueKey, Array(t.Name, t.Start)
                        End If
                    End If
                End If
            End If
        End If
        On Error GoTo 0
    Next t
    
    ' Trier et limiter à 3
    html = html & "<table class='progress-table'>" & vbCrLf
    html = html & "  <thead>" & vbCrLf
    html = html & "    <tr>" & vbCrLf
    html = html & "      <th>Tâche</th>" & vbCrLf
    html = html & "      <th>Date démarrage prévue</th>" & vbCrLf
    html = html & "    </tr>" & vbCrLf
    html = html & "  </thead>" & vbCrLf
    html = html & "  <tbody>" & vbCrLf
    
    If candidateTasks.Count > 0 Then
        ' Trier les clés
        Dim keys() As Variant
        ReDim keys(candidateTasks.Count - 1)
        
        i = 0
        For Each k In candidateTasks.Keys
            keys(i) = k
            i = i + 1
        Next k
        
        ' Tri à bulles
        For i = 0 To UBound(keys) - 1
            For j = i + 1 To UBound(keys)
                If keys(i) > keys(j) Then
                    temp = keys(i)
                    keys(i) = keys(j)
                    keys(j) = temp
                End If
            Next j
        Next i
        
        ' Afficher les 3 premières
        maxRows = 3
        If UBound(keys) + 1 < maxRows Then maxRows = UBound(keys) + 1
        
        For i = 0 To maxRows - 1
            taskName = candidateTasks(keys(i))(0)
            taskStart = candidateTasks(keys(i))(1)
            
            html = html & "    <tr>" & vbCrLf
            html = html & "      <td class='action-name'>" & EncodeHTML(taskName) & "</td>" & vbCrLf
            html = html & "      <td>" & Format(taskStart, "dd/mm/yyyy") & "</td>" & vbCrLf
            html = html & "    </tr>" & vbCrLf
        Next i
    Else
        html = html & "    <tr><td colspan='2' style='text-align:center;color:#666;'>Toutes les tâches sont démarrées</td></tr>" & vbCrLf
    End If
    
    html = html & "  </tbody>" & vbCrLf
    html = html & "</table>" & vbCrLf
    
    BuildClientSection3_ProchainsStarts = html
End Function


' ============================================================================
'  CLIENT SECTION 4 : Planning 3 semaines
' ============================================================================

Private Function BuildClientSection4_AvancementSemaine() As String
    Dim html As String
    Dim t As Task
    Dim planningData As Object
    Dim today As Date
    Dim week1Start As Date
    Dim week4End As Date
    Dim weekStart As Date
    Dim weekNum As Integer
    Dim weekLabel As String
    Dim wNum As Integer
    Dim wStart As Date
    Dim wEndDate As Date
    Dim zone As String
    Dim entreprise As String
    Dim keyZoneEnt As String
    Dim weekKey As String
    Dim a As Assignment
    Dim r As Resource
    Dim rhCount As Long
    Dim currentActivities As String
    Dim keyEntry As Variant
    Dim zoneData As Object
    Dim weekK As String
    Dim rhCnt As Long
    Dim activities As String
    Dim daysFromMonday As Integer
    Dim zoneDict As Object
    Dim weekData As Object
    
    Set planningData = CreateObject("Scripting.Dictionary")
    
    today = Date
    
    ' Trouver le lundi de la semaine courante
    daysFromMonday = Weekday(today, vbMonday) - 1
    week1Start = today - daysFromMonday
    week4End = week1Start + 27 ' 4 semaines (7 jours x 4 - 1)
    
    html = "<h3 style='font-size:14px;font-weight:bold;margin:40px 0 20px 0;'>PLANNING 3 SEMAINES</h3>" & vbCrLf
    html = html & "<table class='progress-table' style='font-size:10px;width:100%;'>" & vbCrLf
    html = html & "  <thead>" & vbCrLf
    html = html & "    <tr>" & vbCrLf
    html = html & "      <th rowspan='2' style='width:80px;'>Zone</th>" & vbCrLf
    html = html & "      <th rowspan='2' style='width:100px;'>Entreprise</th>" & vbCrLf
    
    ' En-têtes des semaines S6, S7, S8, S9
    For weekNum = 1 To 4
        weekStart = week1Start + (weekNum - 1) * 7
        weekLabel = "S" & (5 + weekNum) ' S6, S7, S8, S9
        html = html & "      <th colspan='2' style='background:#d0d0d0;font-weight:bold;'>" & weekLabel & "<br><span style='font-size:9px;font-weight:normal;'>" & Format(weekStart, "dd/mm/yyyy") & "</span></th>" & vbCrLf
    Next weekNum
    
    html = html & "    </tr>" & vbCrLf
    html = html & "    <tr>" & vbCrLf
    For weekNum = 1 To 4
        html = html & "      <th style='width:50px;background:#e8e8e8;'>RH</th>" & vbCrLf
        html = html & "      <th style='min-width:150px;background:#e8e8e8;'>Activité</th>" & vbCrLf
    Next weekNum
    html = html & "    </tr>" & vbCrLf
    html = html & "  </thead>" & vbCrLf
    html = html & "  <tbody>" & vbCrLf
    
    ' Collecter les données par Zone + Entreprise
    For Each t In ActiveProject.Tasks
        On Error Resume Next
        If Not t Is Nothing And Not t.Summary Then
            If Not IsEmpty(t.Start) And Not IsNull(t.Start) And _
               Not IsEmpty(t.Finish) And Not IsNull(t.Finish) Then
                
                ' Vérifier si la tâche est dans la période
                If t.Start <= week4End And t.Finish >= week1Start Then
                    zone = IIf(Len(t.Text3) > 0, CStr(t.Text3), "Sans zone")
                    entreprise = IIf(Len(t.Text1) > 0, CStr(t.Text1), "Non défini")
                    
                    keyZoneEnt = zone & "|" & entreprise
                    
                    If Not planningData.Exists(keyZoneEnt) Then
                        Set zoneDict = CreateObject("Scripting.Dictionary")
                        zoneDict.Add "zone", zone
                        zoneDict.Add "entreprise", entreprise
                        zoneDict.Add "weeks", CreateObject("Scripting.Dictionary")
                        planningData.Add keyZoneEnt, zoneDict
                    End If
                    
                    ' Déterminer dans quelle(s) semaine(s) la tâche se trouve
                    For wNum = 1 To 4
                        wStart = week1Start + (wNum - 1) * 7
                        wEndDate = wStart + 6
                        
                        If t.Start <= wEndDate And t.Finish >= wStart Then
                            weekKey = "S" & (5 + wNum)
                            
                            If Not planningData(keyZoneEnt)("weeks").Exists(weekKey) Then
                                Set weekData = CreateObject("Scripting.Dictionary")
                                weekData.Add "rh", 0
                                weekData.Add "activities", ""
                                planningData(keyZoneEnt)("weeks").Add weekKey, weekData
                            End If
                            
                            ' Compter les RH (ressources de type Travail)
                            rhCount = 0
                            
                            For Each a In t.Assignments
                                If Not a Is Nothing Then
                                    Set r = a.Resource
                                    If Not r Is Nothing Then
                                        If r.Type = pjResourceTypeWork Then
                                            rhCount = rhCount + 1
                                        End If
                                    End If
                                End If
                            Next a
                            
                            ' Mettre à jour le compteur RH (prendre le max si plusieurs tâches)
                            If rhCount > planningData(keyZoneEnt)("weeks")(weekKey)("rh") Then
                                planningData(keyZoneEnt)("weeks")(weekKey)("rh") = rhCount
                            End If
                            
                            ' Ajouter l'activité
                            currentActivities = planningData(keyZoneEnt)("weeks")(weekKey)("activities")
                            If Len(currentActivities) > 0 Then
                                ' Éviter les doublons
                                If InStr(currentActivities, t.Name) = 0 Then
                                    currentActivities = currentActivities & ", " & t.Name
                                End If
                            Else
                                currentActivities = t.Name
                            End If
                            planningData(keyZoneEnt)("weeks")(weekKey)("activities") = currentActivities
                        End If
                    Next wNum
                End If
            End If
        End If
        On Error GoTo 0
    Next t
    
    ' Générer les lignes du tableau
    If planningData.Count > 0 Then
        For Each keyEntry In planningData.Keys
            Set zoneData = planningData(keyEntry)
            
            ' Une seule ligne avec RH et Activité côte à côte
            html = html & "    <tr>" & vbCrLf
            html = html & "      <td class='action-name' style='text-align:center;'>" & EncodeHTML(zoneData("zone")) & "</td>" & vbCrLf
            html = html & "      <td class='action-name' style='text-align:center;'>" & EncodeHTML(zoneData("entreprise")) & "</td>" & vbCrLf
            
            For weekNum = 1 To 4
                weekK = "S" & (5 + weekNum)
                rhCnt = 0
                activities = "Pas d'accès PL"
                
                If zoneData("weeks").Exists(weekK) Then
                    rhCnt = zoneData("weeks")(weekK)("rh")
                    activities = zoneData("weeks")(weekK)("activities")
                    If Len(activities) = 0 Then activities = "Pas d'accès PL"
                End If
                
                html = html & "      <td style='text-align:center;font-weight:bold;'>" & rhCnt & "</td>" & vbCrLf
                html = html & "      <td style='text-align:left;padding:5px;'>" & EncodeHTML(activities) & "</td>" & vbCrLf
            Next weekNum
            
            html = html & "    </tr>" & vbCrLf
        Next keyEntry
    Else
        html = html & "    <tr><td colspan='10' style='text-align:center;color:#666;'>Aucune tâche dans les 4 prochaines semaines</td></tr>" & vbCrLf
    End If
    
    html = html & "  </tbody>" & vbCrLf
    html = html & "</table>" & vbCrLf
    
    BuildClientSection4_AvancementSemaine = html
End Function


' ============================================================================
'  VUE S-CURVE : Courbe en S (Prévu vs Réel)
' ============================================================================

Private Function BuildViewSCurve() As String
    Dim html As String
    Dim t As Task
    Dim dateMin As Date, dateMax As Date, dateToday As Date
    Dim totalWorkPlanned As Double, totalWorkDone As Double
    Dim firstDate As Boolean
    Dim debugLog As String
    Dim dateDict As Object
    Dim metier As String
    Dim taskWork As Double, taskWorkDone As Double
    Dim taskPct As Double
    Dim a As Assignment, r As Resource
    Dim qtyTotal As Double
    Dim currentDate As Date
    Dim taskDuration As Long
    Dim dailyWork As Double
    Dim dateKey As String
    Dim dayData As Object
    Dim actualEndDate As Date
    Dim actualDuration As Long
    Dim dailyActual As Double
    Dim actualDateKey As String
    Dim dayDataNew As Object
    Dim plannedData As String, actualData As String
    Dim cumulPlanned As Double, cumulActual As Double
    Dim sortedDates() As String
    Dim i As Long
    Dim dk As Variant
    Dim j As Long, temp As String
    Dim dkey As String
    Dim pctPlanned As Double, pctActual As Double
    
    dateToday = Date
    firstDate = True
    totalWorkPlanned = 0
    totalWorkDone = 0
    debugLog = "=== DEBUG S-CURVE ===" & vbCrLf
    
    ' Dictionnaire : Date -> {planned: Double, actual: Double}
    Set dateDict = CreateObject("Scripting.Dictionary")
    
    ' Collecter toutes les tâches MECA et leurs données temporelles
    For Each t In ActiveProject.Tasks
        If Not t Is Nothing And Not t.Summary Then
            metier = IIf(Len(t.Text4) > 0, CStr(t.Text4), "")
            
            debugLog = debugLog & "Tache: " & t.Name & " | Text4=" & metier & " | MapGroup=" & MapGroup(metier) & vbCrLf
            
            If MapGroup(metier) = "meca" Then
                debugLog = debugLog & "  -> MECA detectee!" & vbCrLf
                ' Calculer le travail pour cette tâche
                taskWork = 0
                taskWorkDone = 0
                taskPct = 0
                
                ' Récupérer le % d'avancement de la tâche
                If Not IsEmpty(t.PercentComplete) And Not IsNull(t.PercentComplete) Then
                    taskPct = CDbl(t.PercentComplete)
                End If
                
                debugLog = debugLog & "  %Acheve tache=" & taskPct & vbCrLf
                
                ' Parcourir les affectations de ressources Material (consommables)
                On Error Resume Next
                For Each a In t.Assignments
                    If Not a Is Nothing Then
                        Set r = a.Resource
                        If Not r Is Nothing And r.Type = pjResourceTypeMaterial Then
                            qtyTotal = 0
                            
                            If Not IsEmpty(a.Units) And Not IsNull(a.Units) Then
                                qtyTotal = CDbl(a.Units)
                            End If
                            
                            debugLog = debugLog & "    Ressource: " & r.Name & " | Units=" & qtyTotal
                            debugLog = debugLog & vbCrLf
                            
                            taskWork = taskWork + qtyTotal
                        End If
                    End If
                Next a
                On Error GoTo 0
                
                ' Calculer le travail réalisé basé sur le % de la tâche
                If taskWork > 0 And taskPct > 0 Then
                    taskWorkDone = (taskWork * taskPct) / 100#
                End If
                
                debugLog = debugLog & "  Travail: prevu=" & taskWork & " | realise=" & taskWorkDone & vbCrLf
                
                If taskWork > 0 Then
                    totalWorkPlanned = totalWorkPlanned + taskWork
                    totalWorkDone = totalWorkDone + taskWorkDone
                    
                    ' Déterminer les dates min/max du projet
                    If firstDate Then
                        dateMin = t.Start
                        dateMax = t.Finish
                        firstDate = False
                    Else
                        If t.Start < dateMin Then dateMin = t.Start
                        If t.Finish > dateMax Then dateMax = t.Finish
                    End If
                    
                    ' Répartir le travail prévu uniformément entre Start et Finish
                    taskDuration = t.Finish - t.Start
                    If taskDuration <= 0 Then taskDuration = 1
                    
                    dailyWork = taskWork / taskDuration
                    
                    ' Ajouter le travail prévu jour par jour
                    currentDate = t.Start
                    Do While currentDate <= t.Finish
                        dateKey = Format(currentDate, "yyyy-mm-dd")
                        
                        If Not dateDict.Exists(dateKey) Then
                            Set dayData = CreateObject("Scripting.Dictionary")
                            dayData.Add "planned", dailyWork
                            dayData.Add "actual", 0#
                            dateDict.Add dateKey, dayData
                        Else
                            dateDict(dateKey)("planned") = dateDict(dateKey)("planned") + dailyWork
                        End If
                        
                        currentDate = currentDate + 1
                    Loop
                    
                    ' Répartir le travail réalisé proportionnellement sur la durée de la tâche
                    If taskWorkDone > 0 Then
                        ' Déterminer la date de fin de réalisation
                        If t.Finish < dateToday Then
                            actualEndDate = t.Finish
                        Else
                            actualEndDate = dateToday
                        End If
                        
                        ' Calculer la durée réelle
                        actualDuration = actualEndDate - t.Start
                        If actualDuration <= 0 Then actualDuration = 1
                        
                        ' Répartir le travail réalisé jour par jour
                        dailyActual = taskWorkDone / actualDuration
                        
                        currentDate = t.Start
                        Do While currentDate <= actualEndDate
                            actualDateKey = Format(currentDate, "yyyy-mm-dd")
                            
                            If dateDict.Exists(actualDateKey) Then
                                dateDict(actualDateKey)("actual") = dateDict(actualDateKey)("actual") + dailyActual
                            Else
                                ' Créer la date si elle n'existe pas (cas où actualEndDate > t.Finish)
                                Set dayDataNew = CreateObject("Scripting.Dictionary")
                                dayDataNew.Add "planned", 0#
                                dayDataNew.Add "actual", dailyActual
                                dateDict.Add actualDateKey, dayDataNew
                            End If
                            
                            currentDate = currentDate + 1
                        Loop
                    End If
                End If
            End If
        End If
    Next t
    
    ' Construire les données JSON pour Chart.js
    cumulPlanned = 0
    cumulActual = 0
    
    plannedData = "["
    actualData = "["
    
    If dateDict.Count > 0 And totalWorkPlanned > 0 Then
        ' Trier les dates et générer les points cumulatifs
        ReDim sortedDates(dateDict.Count - 1)
        
        i = 0
        For Each dk In dateDict.Keys
            sortedDates(i) = CStr(dk)
            i = i + 1
        Next dk
        
        ' Tri des dates
        For i = 0 To UBound(sortedDates) - 1
            For j = i + 1 To UBound(sortedDates)
                If sortedDates(i) > sortedDates(j) Then
                    temp = sortedDates(i)
                    sortedDates(i) = sortedDates(j)
                    sortedDates(j) = temp
                End If
            Next j
        Next i
        
        ' Générer les points cumulatifs
        For i = 0 To UBound(sortedDates)
            dkey = sortedDates(i)
            
            cumulPlanned = cumulPlanned + dateDict(dkey)("planned")
            cumulActual = cumulActual + dateDict(dkey)("actual")
            
            pctPlanned = (cumulPlanned / totalWorkPlanned) * 100
            pctActual = (cumulActual / totalWorkPlanned) * 100
            
            If pctPlanned > 100 Then pctPlanned = 100
            If pctActual > 100 Then pctActual = 100
            
            If i > 0 Then
                plannedData = plannedData & ","
                actualData = actualData & ","
            End If
            
            plannedData = plannedData & "{x:""" & dkey & """,y:" & Replace(CStr(Round(pctPlanned, 2)), ",", ".") & "}"
            actualData = actualData & "{x:""" & dkey & """,y:" & Replace(CStr(Round(pctActual, 2)), ",", ".") & "}"
        Next i
    End If
    
    plannedData = plannedData & "]"
    actualData = actualData & "]"
    
    debugLog = debugLog & vbCrLf & "TOTAL: prevu=" & totalWorkPlanned & " | realise=" & totalWorkDone & vbCrLf
    debugLog = debugLog & "Dates dict count: " & dateDict.Count & vbCrLf
    
    ' Générer le HTML avec le graphique
    html = "<div class='view-section' data-view='scurve'>" & vbCrLf
    html = html & "  <div style='padding:20px;'>" & vbCrLf
    html = html & "    <h2 style='text-align:center;margin-bottom:20px;'>COURBE EN S - MÉCANIQUE</h2>" & vbCrLf
    html = html & "    <div style='text-align:center;margin-bottom:10px;color:#666;'>" & vbCrLf
    html = html & "      <span style='margin-right:20px;'>📊 Total prévu : " & Round(totalWorkPlanned, 0) & " unités</span>" & vbCrLf
    html = html & "      <span style='margin-right:20px;'>✅ Total réalisé : " & Round(totalWorkDone, 0) & " unités</span>" & vbCrLf
    html = html & "      <span>📈 Avancement : " & Round((totalWorkDone / IIf(totalWorkPlanned > 0, totalWorkPlanned, 1)) * 100, 1) & "%</span>" & vbCrLf
    html = html & "    </div>" & vbCrLf
    html = html & "    <div style='max-width:1200px;margin:0 auto;height:600px;'>" & vbCrLf
    html = html & "      <canvas id='scurveChart'></canvas>" & vbCrLf
    html = html & "    </div>" & vbCrLf
    html = html & "    <div style='margin-top:20px;padding:20px;background:#f5f5f5;border:1px solid #ddd;font-family:monospace;font-size:11px;white-space:pre-wrap;'>" & vbCrLf
    html = html & EncodeHTML(debugLog) & vbCrLf
    html = html & "    </div>" & vbCrLf
    html = html & "    <script>" & vbCrLf
    html = html & "      const plannedData = " & plannedData & ";" & vbCrLf
    html = html & "      const actualData = " & actualData & ";" & vbCrLf
    html = html & "      " & vbCrLf
    html = html & "      document.addEventListener('DOMContentLoaded', function() {" & vbCrLf
    html = html & "        const scurveButton = document.querySelector('[data-view=""scurve""]');" & vbCrLf
    html = html & "        if (scurveButton) {" & vbCrLf
    html = html & "          scurveButton.addEventListener('click', function() {" & vbCrLf
    html = html & "            setTimeout(renderSCurve, 100);" & vbCrLf
    html = html & "          });" & vbCrLf
    html = html & "        }" & vbCrLf
    html = html & "      });" & vbCrLf
    html = html & "      " & vbCrLf
    html = html & "      function renderSCurve() {" & vbCrLf
    html = html & "        const ctx = document.getElementById('scurveChart');" & vbCrLf
    html = html & "        if (!ctx || window.scurveChartInstance) return;" & vbCrLf
    html = html & "        " & vbCrLf
    html = html & "        window.scurveChartInstance = new Chart(ctx, {" & vbCrLf
    html = html & "          type: 'line'," & vbCrLf
    html = html & "          data: {" & vbCrLf
    html = html & "            datasets: [" & vbCrLf
    html = html & "              {" & vbCrLf
    html = html & "                label: 'Prévu (Baseline)'," & vbCrLf
    html = html & "                data: plannedData," & vbCrLf
    html = html & "                borderColor: '#4CAF50'," & vbCrLf
    html = html & "                backgroundColor: 'rgba(76, 175, 80, 0.1)'," & vbCrLf
    html = html & "                borderWidth: 3," & vbCrLf
    html = html & "                borderDash: [5, 5]," & vbCrLf
    html = html & "                tension: 0.4," & vbCrLf
    html = html & "                fill: false," & vbCrLf
    html = html & "                pointRadius: 0" & vbCrLf
    html = html & "              }," & vbCrLf
    html = html & "              {" & vbCrLf
    html = html & "                label: 'Réel (Actuel)'," & vbCrLf
    html = html & "                data: actualData," & vbCrLf
    html = html & "                borderColor: '#FFD700'," & vbCrLf
    html = html & "                backgroundColor: 'rgba(255, 215, 0, 0.1)'," & vbCrLf
    html = html & "                borderWidth: 3," & vbCrLf
    html = html & "                tension: 0.4," & vbCrLf
    html = html & "                fill: false," & vbCrLf
    html = html & "                pointRadius: 3," & vbCrLf
    html = html & "                pointHoverRadius: 6" & vbCrLf
    html = html & "              }" & vbCrLf
    html = html & "            ]" & vbCrLf
    html = html & "          }," & vbCrLf
    html = html & "          options: {" & vbCrLf
    html = html & "            responsive: true," & vbCrLf
    html = html & "            maintainAspectRatio: false," & vbCrLf
    html = html & "            interaction: {" & vbCrLf
    html = html & "              mode: 'index'," & vbCrLf
    html = html & "              intersect: false" & vbCrLf
    html = html & "            }," & vbCrLf
    html = html & "            scales: {" & vbCrLf
    html = html & "              x: {" & vbCrLf
    html = html & "                type: 'time'," & vbCrLf
    html = html & "                time: { unit: 'day', displayFormats: { day: 'DD/MM/YY' } }," & vbCrLf
    html = html & "                title: { display: true, text: 'Date', font: { size: 14, weight: 'bold' } }," & vbCrLf
    html = html & "                grid: { color: 'rgba(0, 0, 0, 0.1)' }" & vbCrLf
    html = html & "              }," & vbCrLf
    html = html & "              y: {" & vbCrLf
    html = html & "                min: 0," & vbCrLf
    html = html & "                max: 100," & vbCrLf
    html = html & "                title: { display: true, text: '% Avancement Cumulé', font: { size: 14, weight: 'bold' } }," & vbCrLf
    html = html & "                grid: { color: 'rgba(0, 0, 0, 0.1)' }," & vbCrLf
    html = html & "                ticks: { callback: function(value) { return value + '%'; } }" & vbCrLf
    html = html & "              }" & vbCrLf
    html = html & "            }," & vbCrLf
    html = html & "            plugins: {" & vbCrLf
    html = html & "              legend: { display: true, position: 'top', labels: { font: { size: 13 } } }," & vbCrLf
    html = html & "              tooltip: {" & vbCrLf
    html = html & "                callbacks: {" & vbCrLf
    html = html & "                  label: function(context) {" & vbCrLf
    html = html & "                    return context.dataset.label + ': ' + context.parsed.y.toFixed(1) + '%';" & vbCrLf
    html = html & "                  }" & vbCrLf
    html = html & "                }" & vbCrLf
    html = html & "              }" & vbCrLf
    html = html & "            }" & vbCrLf
    html = html & "          }" & vbCrLf
    html = html & "        });" & vbCrLf
    html = html & "      }" & vbCrLf
    html = html & "    </script>" & vbCrLf
    html = html & "  </div>" & vbCrLf
    html = html & "</div>" & vbCrLf
    
    BuildViewSCurve = html
End Function


Private Function JsonEscape(ByVal txt As String) As String
    txt = Replace(txt, "\", "\\")
    txt = Replace(txt, """", "\""")
    txt = Replace(txt, vbCrLf, "\n")
    txt = Replace(txt, vbCr, "\n")
    txt = Replace(txt, vbLf, "\n")
    JsonEscape = txt
End Function


' ============================================================================
'  FOOTER : Fermeture HTML
' ============================================================================

Private Function BuildHTMLFooter() As String
    Dim f As String
    f = "</div>" & vbCrLf  ' Fermeture dashboard-container
    f = f & "</body>" & vbCrLf
    f = f & "</html>"
    BuildHTMLFooter = f
End Function


' ============================================================================
'  DEBUG : Fichier log simplifié
' ============================================================================

Private Function CreateDebugLog() As String
    Dim log As String
    Dim t As Task
    Dim zone As String
    Dim metier As String
    Dim allTasksDict As Object
    Dim zoneStatsDict As Object
    Dim globalTotal As Object
    Dim taskKey As String
    Dim taskInfo As Object
    Dim heuresPrevu As Double, heuresActuel As Double
    Dim a As Assignment, r As Resource
    Dim qCalc As Double, pctAssignment As Double
    Dim zoneTotal As Object
    Dim zKey As Variant
    Dim tKey As Variant
    Dim tKey2 As Variant
    Dim nameZoneDict As Object
    Dim taskName As Variant
    Dim tempTaskName As String
    Dim doubleCount As Integer
    Dim sumZonesPrevu As Double, sumZonesActuel As Double
    Dim ecartPrevu As Double, ecartActuel As Double
    Dim zonePct As Double
    Dim summaryCount As Integer
    Dim pctSummary As Double
    
    ' Initialisation des dictionnaires
    Set allTasksDict = CreateObject("Scripting.Dictionary")
    Set zoneStatsDict = CreateObject("Scripting.Dictionary")
    Set globalTotal = CreateObject("Scripting.Dictionary")
    globalTotal.Add "heuresPrevu", 0
    globalTotal.Add "heuresActuel", 0
    globalTotal.Add "taskCount", 0
    
    log = String(80, "=") & vbCrLf
    log = log & "  ANALYSE DIAGNOSTIC COMPLETE - ZONES ET DOUBLONS" & vbCrLf
    log = log & String(80, "=") & vbCrLf & vbCrLf
    log = log & "PROJET : " & ActiveProject.Name & vbCrLf
    log = log & "DATE EXPORT : " & Now() & vbCrLf & vbCrLf
    
    ' ========================================================================
    ' ETAPE 1 : INVENTAIRE COMPLET DES TACHES
    ' ========================================================================
    log = log & vbCrLf & String(80, "=") & vbCrLf
    log = log & "ETAPE 1 : INVENTAIRE COMPLET DES TACHES" & vbCrLf
    log = log & String(80, "=") & vbCrLf & vbCrLf
    
    For Each t In ActiveProject.Tasks
        If Not t Is Nothing And Not t.Summary Then
            zone = IIf(Len(t.Text3) > 0, CStr(t.Text3), "SANS_ZONE")
            metier = IIf(Len(t.Text4) > 0, CStr(t.Text4), "N/A")
            
            heuresPrevu = 0
            heuresActuel = 0
            
            ' Calcul des heures (ressources Material)
            On Error Resume Next
            For Each a In t.Assignments
                If Not a Is Nothing Then
                    Set r = a.Resource
                    If Not r Is Nothing And r.Type = pjResourceTypeMaterial Then
                        qCalc = 0
                        If Not IsEmpty(a.Units) And Not IsNull(a.Units) Then
                            qCalc = CDbl(a.Units)
                        End If
                        heuresPrevu = heuresPrevu + qCalc
                    End If
                End If
            Next a
            On Error GoTo 0
            
            ' Calcul heures actuelles
            pctAssignment = 0
            If Not IsEmpty(t.PercentComplete) And Not IsNull(t.PercentComplete) Then
                pctAssignment = CDbl(t.PercentComplete)
            End If
            heuresActuel = heuresPrevu * pctAssignment / 100#
            
            ' Stockage dans le dictionnaire
            taskKey = "T" & t.UniqueID
            Set taskInfo = CreateObject("Scripting.Dictionary")
            taskInfo.Add "id", t.UniqueID
            taskInfo.Add "name", t.Name
            taskInfo.Add "zone", zone
            taskInfo.Add "metier", metier
            taskInfo.Add "heuresPrevu", heuresPrevu
            taskInfo.Add "heuresActuel", heuresActuel
            taskInfo.Add "pct", pctAssignment
            taskInfo.Add "start", t.Start
            taskInfo.Add "finish", t.Finish
            
            allTasksDict.Add taskKey, taskInfo
            
            ' Mise à jour des totaux globaux
            globalTotal("heuresPrevu") = globalTotal("heuresPrevu") + heuresPrevu
            globalTotal("heuresActuel") = globalTotal("heuresActuel") + heuresActuel
            globalTotal("taskCount") = globalTotal("taskCount") + 1
            
            ' Initialiser les stats par zone si nécessaire
            If Not zoneStatsDict.Exists(zone) Then
                Set zoneTotal = CreateObject("Scripting.Dictionary")
                zoneTotal.Add "heuresPrevu", 0
                zoneTotal.Add "heuresActuel", 0
                zoneTotal.Add "taskCount", 0
                zoneTotal.Add "tasks", CreateObject("Scripting.Dictionary")
                zoneStatsDict.Add zone, zoneTotal
            End If
            
            ' Ajouter la tâche à la zone
            zoneStatsDict(zone)("heuresPrevu") = zoneStatsDict(zone)("heuresPrevu") + heuresPrevu
            zoneStatsDict(zone)("heuresActuel") = zoneStatsDict(zone)("heuresActuel") + heuresActuel
            zoneStatsDict(zone)("taskCount") = zoneStatsDict(zone)("taskCount") + 1
            zoneStatsDict(zone)("tasks").Add taskKey, True
            
            log = log & "ID=" & t.UniqueID & " | " & t.Name & vbCrLf
            log = log & "  Zone: " & zone & " | Metier: " & metier & " | %: " & pctAssignment & "%" & vbCrLf
            log = log & "  Heures: Prevu=" & Format(heuresPrevu, "0.00") & "h | Actuel=" & Format(heuresActuel, "0.00") & "h" & vbCrLf
            log = log & "  Dates: " & Format(t.Start, "dd/mm/yyyy") & " -> " & Format(t.Finish, "dd/mm/yyyy") & vbCrLf
            log = log & vbCrLf
        End If
    Next t
    
    ' ========================================================================
    ' ETAPE 2 : DETECTION DES DOUBLONS
    ' ========================================================================
    log = log & vbCrLf & String(80, "=") & vbCrLf
    log = log & "ETAPE 2 : DETECTION DES DOUBLONS (meme tache dans plusieurs zones)" & vbCrLf
    log = log & String(80, "=") & vbCrLf & vbCrLf
    
    Set nameZoneDict = CreateObject("Scripting.Dictionary")
    
    ' Regrouper par nom de tâche
    For Each tKey In allTasksDict.Keys
        Set taskInfo = allTasksDict(tKey)
        tempTaskName = taskInfo("name")
        
        If Not nameZoneDict.Exists(tempTaskName) Then
            Set nameZoneDict(tempTaskName) = CreateObject("Scripting.Dictionary")
        End If
        
        nameZoneDict(tempTaskName).Add tKey, taskInfo
    Next tKey
    
    ' Identifier les doublons
    doubleCount = 0
    For Each taskName In nameZoneDict.Keys
        If nameZoneDict(taskName).Count > 1 Then
            doubleCount = doubleCount + 1
            log = log & "DOUBLON #" & doubleCount & " : " & taskName & vbCrLf
            log = log & "  Present dans " & nameZoneDict(taskName).Count & " zones :" & vbCrLf
            
            For Each tKey2 In nameZoneDict(taskName).Keys
                Set taskInfo = nameZoneDict(taskName)(tKey2)
                log = log & "    - Zone: " & taskInfo("zone") & " | Heures: " & Format(taskInfo("heuresPrevu"), "0.00") & "h" & vbCrLf
            Next tKey2
            log = log & vbCrLf
        End If
    Next taskName
    
    If doubleCount = 0 Then
        log = log & "OK - Aucun doublon detecte" & vbCrLf
    Else
        log = log & "ALERTE - TOTAL : " & doubleCount & " taches en doublon" & vbCrLf
    End If
    
    ' ========================================================================
    ' ETAPE 3 : STATISTIQUES PAR ZONE
    ' ========================================================================
    log = log & vbCrLf & String(80, "=") & vbCrLf
    log = log & "ETAPE 3 : STATISTIQUES PAR ZONE (filtrage individuel)" & vbCrLf
    log = log & String(80, "=") & vbCrLf & vbCrLf
    
    For Each zKey In zoneStatsDict.Keys
        Set zoneTotal = zoneStatsDict(zKey)
        
        log = log & "+--- ZONE : " & zKey & " " & String(60 - Len(CStr(zKey)), "-") & "+" & vbCrLf
        log = log & "| Nombre de taches : " & zoneTotal("taskCount") & vbCrLf
        log = log & "| Heures prevues   : " & Format(zoneTotal("heuresPrevu"), "#,##0.00") & " h" & vbCrLf
        log = log & "| Heures actuelles : " & Format(zoneTotal("heuresActuel"), "#,##0.00") & " h" & vbCrLf
        
        zonePct = 0
        If zoneTotal("heuresPrevu") > 0 Then
            zonePct = (zoneTotal("heuresActuel") / zoneTotal("heuresPrevu")) * 100
        End If
        log = log & "| Avancement       : " & Format(zonePct, "0.00") & " %" & vbCrLf
        log = log & "+" & String(70, "-") & "+" & vbCrLf & vbCrLf
    Next zKey
    
    ' ========================================================================
    ' ETAPE 4 : COMPARAISON MODE "TOUTES AGREGEES" VS SOMME DES ZONES
    ' ========================================================================
    log = log & vbCrLf & String(80, "=") & vbCrLf
    log = log & "ETAPE 4 : COMPARAISON DES MODES DE FILTRAGE" & vbCrLf
    log = log & String(80, "=") & vbCrLf & vbCrLf
    
    sumZonesPrevu = 0
    sumZonesActuel = 0
    
    For Each zKey In zoneStatsDict.Keys
        Set zoneTotal = zoneStatsDict(zKey)
        sumZonesPrevu = sumZonesPrevu + zoneTotal("heuresPrevu")
        sumZonesActuel = sumZonesActuel + zoneTotal("heuresActuel")
    Next zKey
    
    log = log & "MODE 'TOUTES AGREGEES' (toutes taches uniques) :" & vbCrLf
    log = log & "  Taches    : " & globalTotal("taskCount") & vbCrLf
    log = log & "  Heures P. : " & Format(globalTotal("heuresPrevu"), "#,##0.00") & " h" & vbCrLf
    log = log & "  Heures A. : " & Format(globalTotal("heuresActuel"), "#,##0.00") & " h" & vbCrLf
    log = log & vbCrLf
    
    log = log & "SOMME DES ZONES (addition des filtres individuels) :" & vbCrLf
    log = log & "  Heures P. : " & Format(sumZonesPrevu, "#,##0.00") & " h" & vbCrLf
    log = log & "  Heures A. : " & Format(sumZonesActuel, "#,##0.00") & " h" & vbCrLf
    log = log & vbCrLf
    
    ecartPrevu = sumZonesPrevu - globalTotal("heuresPrevu")
    ecartActuel = sumZonesActuel - globalTotal("heuresActuel")
    
    log = log & "ECART (Somme zones - Toutes agregees) :" & vbCrLf
    log = log & "  Heures P. : " & Format(ecartPrevu, "#,##0.00") & " h"
    If ecartPrevu > 0 Then
        log = log & " ALERTE SURPLUS (doublons comptes plusieurs fois)"
    ElseIf ecartPrevu < 0 Then
        log = log & " ALERTE MANQUE"
    Else
        log = log & " OK"
    End If
    log = log & vbCrLf
    
    log = log & "  Heures A. : " & Format(ecartActuel, "#,##0.00") & " h"
    If ecartActuel > 0 Then
        log = log & " ALERTE SURPLUS (doublons comptes plusieurs fois)"
    ElseIf ecartActuel < 0 Then
        log = log & " ALERTE MANQUE"
    Else
        log = log & " OK"
    End If
    log = log & vbCrLf & vbCrLf
    
    ' ========================================================================
    ' ETAPE 5 : TACHES RECAPITULATIVES
    ' ========================================================================
    log = log & vbCrLf & String(80, "=") & vbCrLf
    log = log & "ETAPE 5 : TACHES RECAPITULATIVES" & vbCrLf
    log = log & String(80, "=") & vbCrLf & vbCrLf
    
    summaryCount = 0
    For Each t In ActiveProject.Tasks
        If Not t Is Nothing And t.Summary Then
            summaryCount = summaryCount + 1
            pctSummary = 0
            If Not IsEmpty(t.PercentComplete) And Not IsNull(t.PercentComplete) Then
                pctSummary = CDbl(t.PercentComplete)
            End If
            metier = IIf(Len(t.Text4) > 0, CStr(t.Text4), "N/A")
            zone = IIf(Len(t.Text3) > 0, CStr(t.Text3), "N/A")
            
            log = log & summaryCount & ". " & t.Name & vbCrLf
            log = log & "   Zone: " & zone & " | Metier: " & metier & " | OutlineLevel: " & t.OutlineLevel & " | %: " & pctSummary & "%" & vbCrLf
            log = log & vbCrLf
        End If
    Next t
    
    ' ========================================================================
    ' ETAPE 6 : COHERENCE ENTRE VUE METIERS ET DASHBOARD CLIENT
    ' ========================================================================
    ' Hypothese : le % affiche dans Vue Metiers est le PercentComplete GLOBAL
    ' d'une tache recapitulative. Si cette recap contient des sous-taches de
    ' PLUSIEURS metiers, le % affiché est "pollue" par les autres metiers.
    ' C'est la cause principale des chiffres illogiques entre les deux vues.
    ' ========================================================================
    log = log & vbCrLf & String(80, "=") & vbCrLf
    log = log & "ETAPE 6 : COHERENCE VUE METIERS vs DASHBOARD CLIENT" & vbCrLf
    log = log & String(80, "=") & vbCrLf & vbCrLf
    log = log & "  >> Text3 = Zone | Text4 = Metier" & vbCrLf & vbCrLf
    
    log = log & LogMetierSection("vrd",  "TRANCHEES  (VRD)")
    log = log & LogMetierSection("meca", "STRUCTURES (MECA)")
    log = log & LogMetierSection("elec", "ELECTRICITE (ELEC)")
    
    ' ========================================================================
    ' CONCLUSION
    ' ========================================================================
    log = log & vbCrLf & String(80, "=") & vbCrLf
    log = log & "DIAGNOSTIC FINAL" & vbCrLf
    log = log & String(80, "=") & vbCrLf & vbCrLf
    
    If ecartPrevu > 0 Or ecartActuel > 0 Then
        log = log & "ALERTE PROBLEME DETECTE : Les doublons sont comptes plusieurs fois dans les filtres de zones." & vbCrLf
        log = log & "   Solution recommandee : Utiliser le mode 'Toutes agregees' pour une vision globale correcte." & vbCrLf
    Else
        log = log & "OK COHERENCE VALIDEE : Pas de doublons, les filtres zones donnent des resultats coherents." & vbCrLf
    End If
    
    log = log & vbCrLf & "Fin du diagnostic." & vbCrLf
    
    CreateDebugLog = log
End Function


' ============================================================================
'  DEBUG HELPER : Analyse d'un metier pour duckdebugging
'
'  Reproduit la logique de BuildMetierSection et expose l'incohérence centrale :
'  Une tache recapitulative peut contenir des sous-taches de PLUSIEURS metiers.
'  Son PercentComplete MSProject est un agregat de TOUS ses metiers.
'  => Le % affiche dans Vue Metiers n'est pas le vrai % du metier seul.
'  => C'est la source des chiffres divergents entre Vue Metiers et Dashboard Client.
' ============================================================================

Private Function LogMetierSection(ByVal metierType As String, ByVal metierLabel As String) As String
    Dim log As String
    Dim t As Task, sub1 As Task
    Dim hasMatchingLot As Boolean
    Dim taskZone As String, subLot As String
    Dim pct As Double
    Dim foundList As Object
    Dim key As Variant
    Dim tName As String, tPct As Double, tZone As String
    Dim tLevel As Integer, tOutline As String
    Dim subCountMatch As Long, subCountOther As Long, subCountNone As Long
    Dim sumPctMatch As Double, pctPur As Double
    Dim otherMetiers As String, subPct As Double
    Dim isMatch As Boolean, totalSub As Long
    Dim alertMsg As String
    
    Set foundList = CreateObject("Scripting.Dictionary")
    
    log = "+--- " & metierLabel & " " & String(74 - Len(metierLabel), "-") & vbCrLf
    
    ' Reproduire la même logique de détection que BuildMetierSection
    For Each t In ActiveProject.Tasks
        If Not t Is Nothing And t.Summary And t.OutlineLevel > 1 Then
            hasMatchingLot = False
            taskZone = "SANS_ZONE"
            
            For Each sub1 In ActiveProject.Tasks
                If Not sub1 Is Nothing And Not sub1.Summary Then
                    If sub1.OutlineLevel > t.OutlineLevel And sub1.ID > t.ID And _
                       (sub1.OutlineParent Is Nothing Or sub1.OutlineParent.ID = t.ID Or _
                        IsSubTaskOf(sub1, t)) Then
                        subLot = LCase$(Trim$(IIf(Len(sub1.Text4) > 0, CStr(sub1.Text4), "")))
                        If (metierType = "vrd"  And InStr(subLot, "vrd") > 0) Or _
                           (metierType = "meca" And (InStr(subLot, "meca") > 0 Or InStr(subLot, "mecanique") > 0)) Or _
                           (metierType = "elec" And (InStr(subLot, "elec") > 0 Or InStr(subLot, "electrique") > 0)) Then
                            hasMatchingLot = True
                            If taskZone = "SANS_ZONE" And Len(sub1.Text3) > 0 Then taskZone = CStr(sub1.Text3)
                        End If
                    End If
                End If
            Next sub1
            
            If hasMatchingLot And Not foundList.Exists(t.ID) Then
                pct = 0
                If Not IsEmpty(t.PercentComplete) And Not IsNull(t.PercentComplete) Then pct = CDbl(t.PercentComplete)
                foundList.Add t.ID, Array(t.Name, pct, taskZone, t.OutlineLevel, t.OutlineNumber)
            End If
        End If
    Next t
    
    log = log & "| Recap detectees pour ce metier : " & foundList.Count & vbCrLf
    If foundList.Count = 0 Then
        log = log & "| => AUCUNE - verifier que Text4 des sous-taches contient '" & metierType & "'" & vbCrLf
        log = log & "+" & String(79, "-") & vbCrLf & vbCrLf
        LogMetierSection = log
        Exit Function
    End If
    log = log & "|" & vbCrLf
    
    For Each key In foundList.Keys
        tName    = foundList(key)(0)
        tPct     = foundList(key)(1)
        tZone    = foundList(key)(2)
        tLevel   = foundList(key)(3)
        tOutline = foundList(key)(4)
        
        log = log & "| [ID=" & key & " Niv=" & tLevel & " Plan=" & tOutline & "] " & tName & vbCrLf
        log = log & "|   Zone : " & tZone & vbCrLf
        log = log & "|   % AFFICHE dans Vue Metiers  = " & Format(tPct, "0.0") & _
                    "% (PercentComplete GLOBAL MSProject - tous metiers confondus)" & vbCrLf
        
        ' ---- Décomposer les sous-tâches par métier ----
        subCountMatch = 0 : subCountOther = 0 : subCountNone = 0
        sumPctMatch = 0 : otherMetiers = ""
        
        For Each sub1 In ActiveProject.Tasks
            If Not sub1 Is Nothing And Not sub1.Summary Then
                If IsSubTaskOf(sub1, ActiveProject.Tasks(CLng(key))) Then
                    subPct = 0
                    If Not IsEmpty(sub1.PercentComplete) And Not IsNull(sub1.PercentComplete) Then
                        subPct = CDbl(sub1.PercentComplete)
                    End If
                    subLot = LCase$(Trim$(IIf(Len(sub1.Text4) > 0, CStr(sub1.Text4), "")))
                    isMatch = False
                    If (metierType = "vrd"  And InStr(subLot, "vrd") > 0) Then isMatch = True
                    If (metierType = "meca" And (InStr(subLot, "meca") > 0 Or InStr(subLot, "mecanique") > 0)) Then isMatch = True
                    If (metierType = "elec" And (InStr(subLot, "elec") > 0 Or InStr(subLot, "electrique") > 0)) Then isMatch = True
                    
                    If isMatch Then
                        subCountMatch = subCountMatch + 1
                        sumPctMatch = sumPctMatch + subPct
                    ElseIf Len(subLot) = 0 Then
                        subCountNone = subCountNone + 1
                    Else
                        subCountOther = subCountOther + 1
                        If InStr(otherMetiers, sub1.Text4) = 0 Then
                            otherMetiers = otherMetiers & IIf(Len(otherMetiers) > 0, " + ", "") & sub1.Text4
                        End If
                    End If
                End If
            End If
        Next sub1
        
        totalSub = subCountMatch + subCountOther + subCountNone
        log = log & "|   Sous-taches : " & totalSub & " total"
        log = log & " | " & subCountMatch & " " & UCase(metierType)
        If subCountOther > 0 Then log = log & " | " & subCountOther & " AUTRE (" & otherMetiers & ")"
        If subCountNone  > 0 Then log = log & " | " & subCountNone & " sans Text4"
        log = log & vbCrLf
        
        alertMsg = ""
        If subCountOther > 0 Then
            pctPur = IIf(subCountMatch > 0, sumPctMatch / subCountMatch, 0)
            log = log & "|   % pur " & UCase(metierType) & " (sous-taches " & metierType & " seulement) = " & Format(pctPur, "0.0") & "%" & vbCrLf
            alertMsg = "INCOHERENCE : recap mixte (" & otherMetiers & " inclus)."
            alertMsg = alertMsg & " Affiche=" & Format(tPct, "0.0") & "% vs Pur-" & UCase(metierType) & "=" & Format(pctPur, "0.0") & "%."
            alertMsg = alertMsg & " Ecart=" & Format(Abs(tPct - pctPur), "0.0") & "%."
        ElseIf totalSub = 0 Then
            alertMsg = "ATTENTION : aucune sous-tache directe trouvee (verifier niveaux de plan OutlineNumber)."
        End If
        
        If Len(alertMsg) > 0 Then
            log = log & "|   >>> " & alertMsg & vbCrLf
        Else
            log = log & "|   >>> OK : toutes les sous-taches sont du metier " & UCase(metierType) & ", % coherent." & vbCrLf
        End If
        log = log & "|" & vbCrLf
    Next key
    
    log = log & "+" & String(79, "-") & vbCrLf & vbCrLf
    LogMetierSection = log
End Function


' ============================================================================
'  UTILITAIRES
' ============================================================================

Private Function MapGroup(ByVal s As String) As String
    Dim k As String
    k = LCase$(Trim$(s))
    
    If InStr(k, "elec") > 0 Or InStr(k, "electrique") > 0 Then
        MapGroup = "elec"
    ElseIf InStr(k, "mec") > 0 Or InStr(k, "mecanique") > 0 Then
        MapGroup = "meca"
    Else
        MapGroup = ""
    End If
End Function

Private Function EncodeHTML(rawText As String) As String
    Dim tmp As String
    tmp = rawText
    tmp = Replace(tmp, "&", "&amp;")
    tmp = Replace(tmp, "<", "&lt;")
    tmp = Replace(tmp, ">", "&gt;")
    tmp = Replace(tmp, """", "&quot;")
    tmp = Replace(tmp, "'", "&#39;")
    EncodeHTML = tmp
End Function
