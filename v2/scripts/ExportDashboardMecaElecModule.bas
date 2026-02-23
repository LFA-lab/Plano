Attribute VB_Name = "ExportDashboardMecaElecModule"
Option Explicit

' ============================================================================
'  CONFIGURATION : dossier de sortie des fichiers générés
'  1) Essaie OUTPUT_DIR (WSL UNC)
'  2) Sinon: dossier du projet actif
'  3) Sinon: Mes Documents utilisateur
' ============================================================================
Private Const OUTPUT_DIR As String = "\\wsl.localhost\Ubuntu\home\ntoi\LFA-lab\Plano\macros\export\"

' ============================================================================
'  MODULE : ExportDashboardMecaElec (HTML)
'  OBJET  : Exporter un dashboard HTML depuis MS Project (Vue Métiers, Client, S-Curve)
'  FORMAT : HTML encodé UTF-8 (écriture via ADODB.Stream)
' ============================================================================

Public Sub ExportDashboardMecaElec()
    On Error GoTo ErrorHandler

    If ActiveProject Is Nothing Then
        MsgBox "Aucun projet actif !", vbCritical
        Exit Sub
    End If

    Dim pj As MSProject.Project
    Set pj = ActiveProject

    ' Titre du dashboard: nom projet ou premier nom de tâche non vide
    Dim dashboardTitle As String
    dashboardTitle = pj.Name
    Dim t As Task
    For Each t In pj.Tasks
        If Not t Is Nothing Then
            If Len(Trim$(t.Name)) > 0 Then
                dashboardTitle = t.Name
                Exit For
            End If
        End If
    Next t

    ' Résoudre le dossier de sortie
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim outDir As String
    If FolderExistsSafe(fso, OUTPUT_DIR) Then
        outDir = OUTPUT_DIR
    ElseIf Len(pj.path) > 0 Then
        outDir = pj.path
    Else
        outDir = Environ$("USERPROFILE") & "\Documents"
    End If
    If Right$(outDir, 1) <> "\" Then outDir = outDir & "\"

    ' Chemins
    Dim ts As String: ts = Format(Now, "yyyymmdd_hhnnss")
    Dim filePath As String: filePath = outDir & "dashboard_mecaelec_" & ts & ".html"
    Dim logPath As String:  logPath = outDir & "dashboard_mecaelec_DEBUG_" & ts & ".txt"

    ' Construire HTML + LOG
    Dim htmlContent As String
    htmlContent = BuildCompleteHTML(dashboardTitle)

    If Len(htmlContent) = 0 Then
        MsgBox "ERREUR : Le contenu HTML est vide !", vbCritical
        Exit Sub
    End If

    Dim logContent As String
    logContent = CreateDebugLog()

    ' Écrire UTF-8
    WriteTextUtf8 filePath, htmlContent
    WriteTextUtf8 logPath, logContent

    MsgBox "Export terminé !" & vbCrLf & vbCrLf & _
           "HTML : " & filePath & vbCrLf & _
           "LOG  : " & logPath, vbInformation, "Export Dashboard"
    Exit Sub

ErrorHandler:
    MsgBox "ERREUR #" & Err.Number & " :" & vbCrLf & Err.Description, vbCritical, "Erreur"
End Sub

' ========================= File I/O (UTF-8) =========================
Private Sub WriteTextUtf8(ByVal path As String, ByVal content As String)
    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    With stm
        .Type = 2                 ' text
        .Charset = "utf-8"
        .Open
        .WriteText content
        .SaveToFile path, 2       ' adSaveCreateOverWrite
        .Close
    End With
End Sub

Private Function FolderExistsSafe(fso As Object, ByVal p As String) As Boolean
    On Error Resume Next
    FolderExistsSafe = fso.FolderExists(p)
    On Error GoTo 0
End Function

' ============================ HTML BUILD ============================

Private Function BuildCompleteHTML(ByVal projectName As String) As String
    Dim html As String
    html = BuildHTMLHeader(projectName)
    html = html & BuildHTMLBody()
    html = html & BuildJavaScript()
    html = html & BuildHTMLFooter()
    BuildCompleteHTML = html
End Function

Private Function BuildHTMLHeader(ByVal projectName As String) As String
    Dim parts As String

    ' Start HTML document
    AddLine parts, "<!DOCTYPE html>"
    AddLine parts, "<html lang='fr'>"
    AddLine parts, "<head>"
    AddLine parts, "  <meta charset='UTF-8'>"
    AddLine parts, "  <meta name='viewport' content='width=device-width, initial-scale=1.0'>"
    AddLine parts, "  <title>Dashboard Interne - " & EncodeHTML(projectName) & "</title>"

    ' Chart.js CDN (required for S-Curve)
    AddLine parts, "  <script src='https://cdn.jsdelivr.net/npm/chart.js@4.4.1/dist/chart.umd.min.js'></script>"
    AddLine parts, "  <script src='https://cdn.jsdelivr.net/npm/chartjs-adapter-date-fns@3'></script>"

    ' Include CSS (already safe in its own function)
    parts = parts & BuildCSS()

    ' Body start
    AddLine parts, "</head>"
    AddLine parts, "<body>"
    AddLine parts, "<div class='dashboard-container'>"
    AddLine parts, "  <h1>Dashboard Interne - " & EncodeHTML(projectName) & "</h1>"

    BuildHTMLHeader = parts
End Function

Private Function BuildCSS() As String
    Dim css As String

    css = "<style>" & vbCrLf

    css = css & Join(Array( _
        "*{margin:0;padding:0;box-sizing:border-box}", _
        "body{font-family:'Segoe UI',Tahoma,Geneva,Verdana,sans-serif;background:#f5f5f5;padding:20px}", _
        ".dashboard-container{max-width:1400px;margin:0 auto;background:#fff;border:2px solid #000;padding:20px}", _
        "h1{font-size:16px;font-weight:700;margin-bottom:20px;text-align:left}" _
    ), vbCrLf) & vbCrLf

    css = css & Join(Array( _
        ".tabs-row{display:flex;gap:8px;flex-wrap:wrap;margin-bottom:15px}", _
        ".tab-button{padding:8px 16px;background:#fff;border:1px solid #000;cursor:pointer;font-size:11px;font-weight:600;transition:.2s}", _
        ".tab-button:hover{background:#e8e8e8}", _
        ".tab-button.active{background:#d0d0d0;border-width:2px}", _
        ".view-section{display:none}", _
        ".view-section.active{display:block}" _
    ), vbCrLf) & vbCrLf

    css = css & Join(Array( _
        ".mechanical-progress-view{padding:20px;border:1px solid #000}", _
        ".mechanical-progress-title{text-align:center;font-size:16px;font-weight:700;margin-bottom:30px}", _
        ".mechanical-progress-grid{max-width:900px;margin:0 auto}", _
        ".mechanical-progress-row{display:grid;grid-template-columns:180px 1fr 80px;align-items:center;gap:10px;margin-bottom:8px}", _
        ".mechanical-progress-label{font-size:9px;font-weight:600;text-align:left;line-height:1.2}", _
        ".mechanical-progress-bar-container{background:#fff;border:2px solid #000;height:16px;position:relative}", _
        ".mechanical-progress-bar{height:100%;background:#f4b400;border-right:1px solid #000;transition:width .3s ease}", _
        ".mechanical-progress-percentage{font-size:9px;font-weight:700;text-align:right}", _
        ".mechanical-progress-row.general{margin-top:15px;padding-top:12px;border-top:2px solid #000}", _
        ".mechanical-progress-row.general .mechanical-progress-label{font-size:10px;font-weight:700}", _
        ".mechanical-progress-row.general .mechanical-progress-percentage{font-size:10px}" _
    ), vbCrLf) & vbCrLf

    css = css & Join(Array( _
        ".electrical-view{padding:20px}", _
        ".progress-table{width:100%;border-collapse:collapse;border:1px solid #000;margin-bottom:20px}", _
        ".progress-table th,.progress-table td{border:1px solid #000;padding:8px;text-align:center;font-size:11px}", _
        ".progress-table th{background:#e0e0e0;font-weight:600}", _
        ".progress-table td.action-name{text-align:left;font-weight:600;background:#f5f5f5}", _
        ".progress-table td.percentage{font-weight:600}", _
        ".progress-table td.percentage.low{color:#d32f2f}", _
        ".progress-table td.percentage.medium{color:#f57c00}", _
        ".progress-table td.percentage.high{color:#388e3c}" _
    ), vbCrLf) & vbCrLf

    css = css & "</style>" & vbCrLf
    BuildCSS = css
End Function

' ============================ BODY (TABS) ===========================

Private Function BuildHTMLBody() As String
    Dim html As String

    ' Onglets
    html = Join(Array( _
        "<div class='tabs-row'>", _
        "  <button class='tab-button active' data-view='metiers'>Vue Métiers</button>", _
        "  <button class='tab-button' data-view='client'>Dashboard Client</button>", _
        "  <button class='tab-button' data-view='scurve'>Courbe S (MECA)</button>", _
        "</div>" _
    ), vbCrLf) & vbCrLf

    ' Vue Métiers (avec sélecteur de sous-zone)
    html = html & BuildViewMetiers()

    ' Vue Client
    html = html & BuildViewClient()

    ' Vue Courbe en S (MECA)
    html = html & BuildViewSCurve()

    BuildHTMLBody = html
End Function

' Append one short line without using line continuations
Private Sub AddLine(ByRef sb As String, ByVal line As String)
    sb = sb & line & vbCrLf
End Sub

Private Function BuildZoneSelector(Optional ByVal selectId As String = "zoneFilter", _
                                   Optional ByVal infoId As String = "zoneInfo") As String
    Dim html As String
    Dim t As Task, zones As Object
    Dim zone As String, zoneKey As Variant
    Dim sortedZones() As String
    Dim i As Long, j As Long, temp As String

    Set zones = CreateObject("Scripting.Dictionary")

    ' Collect distinct Sous-Zones from Text3 (leaf tasks only)
    For Each t In ActiveProject.Tasks
        If Not t Is Nothing And Not t.Summary Then
            zone = IIf(Len(t.Text3) > 0, CStr(t.Text3), "")
            If Len(zone) > 0 And Not zones.Exists(zone) Then zones.Add zone, True
        End If
    Next t

    ' Sort alphabetically
    If zones.Count > 0 Then
        ReDim sortedZones(zones.Count - 1)
        i = 0
        For Each zoneKey In zones.keys
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

    AddLine html, "<div style='margin:15px 0;padding:12px;border:2px solid #000;background:#f9f9f9;'>"
    AddLine html, "  <label style='font-size:11px;font-weight:600;margin-right:10px;'>Affichage par sous-zone :</label>"
    AddLine html, "  <select id='" & selectId & "' style='padding:6px 12px;font-size:11px;border:2px solid #000;background:white;cursor:pointer;min-width:200px;'>"
    AddLine html, "    <option value='all' selected>&#10003; Toutes les zones (consolidé)</option>"

    If zones.Count > 0 Then
        For i = 0 To UBound(sortedZones)
            AddLine html, "    <option value='" & EncodeHTML(sortedZones(i)) & "'>" & EncodeHTML(sortedZones(i)) & "</option>"
        Next i
    End If

    AddLine html, "  </select>"
    AddLine html, "  <span id='" & infoId & "' style='margin-left:15px;font-size:10px;color:#666;'></span>"
    AddLine html, "</div>"

    BuildZoneSelector = html
End Function

Private Function BuildViewMetiers() As String
    Dim html As String

    AddLine html, "<div class='view-section active' data-view='metiers'>"
    ' Zone selector inside the Métiers view
    html = html & BuildZoneSelector()

    AddLine html, "  <div style='display:grid;grid-template-columns:repeat(3,1fr);gap:15px;padding:15px;'>"

    ' VRD
    AddLine html, "    <div style='border:2px solid #000;padding:12px;background:#fff;'>"
    AddLine html, "      <h3 style='text-align:center;font-weight:bold;margin-bottom:15px;font-size:13px;'>TRANCHÉES</h3>"
    html = html & BuildMetierSection("vrd")
    AddLine html, "    </div>"

    ' MECA
    AddLine html, "    <div style='border:2px solid #000;padding:12px;background:#fff;'>"
    AddLine html, "      <h3 style='text-align:center;font-weight:bold;margin-bottom:15px;font-size:13px;'>STRUCTURES</h3>"
    html = html & BuildMetierSection("meca")
    AddLine html, "    </div>"

    ' ELEC
    AddLine html, "    <div style='border:2px solid #000;padding:12px;background:#fff;'>"
    AddLine html, "      <h3 style='text-align:center;font-weight:bold;margin-bottom:15px;font-size:13px;'>ÉLECTRICITÉ</h3>"
    html = html & BuildMetierSection("elec")
    AddLine html, "    </div>"

    AddLine html, "  </div>"
    AddLine html, "</div>"

    BuildViewMetiers = html
End Function

Private Function BuildJavaScript() As String
    Dim js As String

    AddLine js, "<script>"
    AddLine js, "document.addEventListener('DOMContentLoaded', function() {"
    AddLine js, "  const buttons = document.querySelectorAll('.tab-button');"
    AddLine js, "  const views = document.querySelectorAll('.view-section');"
    AddLine js, "  const zoneFilter = document.getElementById('zoneFilter');"
    AddLine js, "  const zoneInfo = document.getElementById('zoneInfo');"

    ' Tabs handling
    AddLine js, "  buttons.forEach(btn => {"
    AddLine js, "    btn.addEventListener('click', function() {"
    AddLine js, "      const targetView = this.getAttribute('data-view');"
    AddLine js, "      buttons.forEach(b => b.classList.remove('active'));"
    AddLine js, "      this.classList.add('active');"
    AddLine js, "      views.forEach(v => {"
    AddLine js, "        v.classList.toggle('active', v.getAttribute('data-view') === targetView);"
    AddLine js, "      });"
    AddLine js, "      if (targetView === 'scurve') { setTimeout(renderSCurve, 100); }"
    AddLine js, "    });"
    AddLine js, "  });"

    ' Zone filter (Métiers)
    AddLine js, "  if (zoneFilter) {"
    AddLine js, "    zoneFilter.addEventListener('change', function() {"
    AddLine js, "      const selectedZone = this.value;"
    AddLine js, "      const rows = document.querySelectorAll('[data-view=""metiers""] [data-zone]');"
    AddLine js, "      let visibleCount = 0, totalCount = 0, sumPct = 0, visibleDataRows = 0;"
    AddLine js, "      rows.forEach(row => {"
    AddLine js, "        const rowZone = row.getAttribute('data-zone');"
    AddLine js, "        const rowPct = parseFloat(row.getAttribute('data-pct')) || 0;"
    AddLine js, "        if (rowZone !== '__GENERAL__') totalCount++;"
    AddLine js, "        if (selectedZone === 'all') {"
    AddLine js, "          row.style.display='';"
    AddLine js, "          visibleCount++;"
    AddLine js, "          if (rowZone !== '__GENERAL__') { sumPct += rowPct; visibleDataRows++; }"
    AddLine js, "        } else {"
    AddLine js, "          if (rowZone === selectedZone) {"
    AddLine js, "            row.style.display='';"
    AddLine js, "            visibleCount++;"
    AddLine js, "            sumPct += rowPct;"
    AddLine js, "            visibleDataRows++;"
    AddLine js, "          } else if (rowZone === '__GENERAL__') {"
    AddLine js, "            row.style.display='none';"
    AddLine js, "          } else {"
    AddLine js, "            row.style.display='none';"
    AddLine js, "          }"
    AddLine js, "        }"
    AddLine js, "      });"

    ' Recompute general row
    AddLine js, "      const generalRows = document.querySelectorAll('[data-view=""metiers""] .mechanical-progress-row.general');"
    AddLine js, "      generalRows.forEach(generalRow => {"
    AddLine js, "        if (selectedZone !== 'all') {"
    AddLine js, "          const avgPct = visibleDataRows > 0 ? sumPct / visibleDataRows : 0;"
    AddLine js, "          const bar = generalRow.querySelector('.mechanical-progress-bar');"
    AddLine js, "          const percentage = generalRow.querySelector('.mechanical-progress-percentage');"
    AddLine js, "          const label = generalRow.querySelector('.mechanical-progress-label');"
    AddLine js, "          if (bar) bar.style.width = avgPct + '%';"
    AddLine js, "          if (percentage) percentage.textContent = avgPct.toFixed(1).replace('.', ',') + '%';"
    AddLine js, "          if (label) label.textContent = 'Avancement ' + selectedZone;"
    AddLine js, "          generalRow.style.display = '';"
    AddLine js, "          generalRow.setAttribute('data-zone', selectedZone);"
    AddLine js, "        } else {"
    AddLine js, "          const origPct = parseFloat(generalRow.getAttribute('data-pct')) || 0;"
    AddLine js, "          generalRow.setAttribute('data-zone', '__GENERAL__');"
    AddLine js, "          generalRow.style.display = '';"
    AddLine js, "          const bar = generalRow.querySelector('.mechanical-progress-bar');"
    AddLine js, "          const percentage = generalRow.querySelector('.mechanical-progress-percentage');"
    AddLine js, "          const label = generalRow.querySelector('.mechanical-progress-label');"
    AddLine js, "          if (bar) bar.style.width = origPct + '%';"
    AddLine js, "          if (percentage) percentage.textContent = origPct.toFixed(1).replace('.', ',') + '%';"
    AddLine js, "          if (label) label.textContent = 'Avancement général';"
    AddLine js, "        }"
    AddLine js, "      });"

    ' Update info text
    AddLine js, "      if (zoneInfo) {"
    AddLine js, "        if (selectedZone === 'all') {"
    AddLine js, "          const totalRows = document.querySelectorAll('[data-view=""metiers""] [data-zone]:not([data-zone=""__GENERAL__""])').length;"
    AddLine js, "          zoneInfo.textContent = totalRows + ' ligne' + (totalRows>1?'s':'') + ' (toutes zones agrégées)';"
    AddLine js, "        } else {"
    AddLine js, "          zoneInfo.textContent = visibleDataRows + ' ligne' + (visibleDataRows>1?'s':'') + ' affichée' + (visibleDataRows>1?'s':'');"
    AddLine js, "        }"
    AddLine js, "      }"
    AddLine js, "    });"
    AddLine js, "    const totalRows = document.querySelectorAll('[data-view=""metiers""] [data-zone]:not([data-zone=""__GENERAL__""])').length;"
    AddLine js, "    if (zoneInfo) zoneInfo.textContent = totalRows + ' ligne' + (totalRows>1?'s':'') + ' (toutes zones agrégées)';"
    AddLine js, "  }"

    ' Zone filter (Client)
    AddLine js, "  const zoneFilterClient = document.getElementById('zoneFilterClient');"
    AddLine js, "  const zoneInfoClient   = document.getElementById('zoneInfoClient');"
    AddLine js, "  if (zoneFilterClient) {"
    AddLine js, "    zoneFilterClient.addEventListener('change', function() {"
    AddLine js, "      const selectedZone = this.value;"
    AddLine js, "      const rows = document.querySelectorAll('[data-view=""client""] [data-zone]');"
    AddLine js, "      let visible = 0;"
    AddLine js, "      rows.forEach(row => {"
    AddLine js, "        const rz = row.getAttribute('data-zone');"
    AddLine js, "        if (selectedZone === 'all' || rz === selectedZone) { row.style.display=''; visible++; }"
    AddLine js, "        else { row.style.display='none'; }"
    AddLine js, "      });"
    AddLine js, "      if (zoneInfoClient) {"
    AddLine js, "        if (selectedZone === 'all') zoneInfoClient.textContent = rows.length + ' lot' + (rows.length>1?'s':'') + ' (toutes zones)';"
    AddLine js, "        else zoneInfoClient.textContent = visible + ' lot' + (visible>1?'s':'') + ' affiché' + (visible>1?'s':'');"
    AddLine js, "      }"
    AddLine js, "    });"
    AddLine js, "    const totalClient = document.querySelectorAll('[data-view=""client""] [data-zone]').length;"
    AddLine js, "    if (zoneInfoClient) zoneInfoClient.textContent = totalClient + ' lot' + (totalClient>1?'s':'') + ' (toutes zones)';"
    AddLine js, "  }"

    AddLine js, "});"
    AddLine js, "</script>"

    BuildJavaScript = js
End Function
Private Function BuildMetierSection(ByVal metierType As String) As String
    Dim html As String
    Dim t As Task
    Dim lot As String, subZone As String
    Dim dur As Double, pct As Double
    Dim avgPct As Double, pctStr As String
    Dim genSumPctDur As Double, genSumDur As Double, genSumPct As Double
    Dim genCount As Long
    Dim hasRows As Boolean

    hasRows = False

    ' Scan all leaf tasks
    For Each t In ActiveProject.Tasks
        If Not t Is Nothing And Not t.Summary Then

            ' Check métier match (Text4)
            lot = LCase$(Trim$(IIf(Len(t.Text4) > 0, CStr(t.Text4), "")))
            If MetierMatch(metierType, lot) Then

                ' Sous-zone (Text3)
                subZone = IIf(Len(t.Text3) > 0, CStr(t.Text3), "SANS_ZONE")

                ' Duration
                dur = 0
                On Error Resume Next
                If Not IsEmpty(t.Duration) And Not IsNull(t.Duration) Then dur = CDbl(t.Duration)
                On Error GoTo 0

                ' % Complete
                pct = 0
                On Error Resume Next
                If Not IsEmpty(t.PercentComplete) And Not IsNull(t.PercentComplete) Then pct = CDbl(t.PercentComplete)
                On Error GoTo 0

                ' Accumulate for general average
                genSumPctDur = genSumPctDur + (pct * dur)
                genSumDur = genSumDur + dur
                genSumPct = genSumPct + pct
                genCount = genCount + 1

                ' Render HTML line
                pctStr = Replace(CStr(pct), ",", ".")
                html = html & _
                    "      <div class='mechanical-progress-row' data-zone='" & EncodeHTML(subZone) & "' data-pct='" & pctStr & "'>" & vbCrLf & _
                    "        <div class='mechanical-progress-label'>" & EncodeHTML(t.Name) & "</div>" & vbCrLf & _
                    "        <div class='mechanical-progress-bar-container'>" & vbCrLf & _
                    "          <div class='mechanical-progress-bar' style='width:" & pctStr & "%;'></div>" & vbCrLf & _
                    "        </div>" & vbCrLf & _
                    "        <div class='mechanical-progress-percentage'>" & Replace(CStr(Round(pct, 1)), ".", ",") & "%</div>" & vbCrLf & _
                    "      </div>" & vbCrLf

                hasRows = True
            End If
        End If
    Next t

    ' Add general average row
    If hasRows Then
        If genSumDur > 0 Then
            avgPct = genSumPctDur / genSumDur
        ElseIf genCount > 0 Then
            avgPct = genSumPct / genCount
        Else
            avgPct = 0
        End If

        pctStr = Replace(CStr(avgPct), ",", ".")

        html = html & _
            "      <div class='mechanical-progress-row general' data-zone='__GENERAL__' data-pct='" & pctStr & "'>" & vbCrLf & _
            "        <div class='mechanical-progress-label'>Avancement général</div>" & vbCrLf & _
            "        <div class='mechanical-progress-bar-container'>" & vbCrLf & _
            "          <div class='mechanical-progress-bar' style='width:" & pctStr & "%;'></div>" & vbCrLf & _
            "        </div>" & vbCrLf & _
            "        <div class='mechanical-progress-percentage'>" & Replace(CStr(Round(avgPct, 1)), ".", ",") & "%</div>" & vbCrLf & _
            "      </div>" & vbCrLf

    Else
        html = html & "      <p style='text-align:center;color:#999;font-size:9px;padding:12px;'>Aucune donnée</p>" & vbCrLf
    End If

    BuildMetierSection = html
End Function

Private Function MetierMatch(ByVal metierType As String, ByVal lotLower As String) As Boolean
    ' lotLower is expected to be already lower-cased (we still guard just in case)
    Dim m As String
    m = LCase$(Trim$(metierType))
    lotLower = LCase$(Trim$(lotLower))

    Select Case m
        Case "vrd"
            MetierMatch = (InStr(lotLower, "vrd") > 0)
        Case "meca"
            MetierMatch = (InStr(lotLower, "meca") > 0 Or InStr(lotLower, "mécanique") > 0 Or InStr(lotLower, "mecanique") > 0)
        Case "elec"
            MetierMatch = (InStr(lotLower, "elec") > 0 Or InStr(lotLower, "électrique") > 0 Or InStr(lotLower, "electrique") > 0)
        Case Else
            MetierMatch = False
    End Select
End Function
Private Function IsSubTaskOf(ByVal childTask As Task, ByVal parentTask As Task) As Boolean
    Dim parentOutline As String
    Dim childOutline As String

    IsSubTaskOf = False

    ' Guard rails
    If childTask Is Nothing Then Exit Function
    If parentTask Is Nothing Then Exit Function

    ' A child must be deeper in outline level
    If childTask.OutlineLevel <= parentTask.OutlineLevel Then Exit Function

    ' Compare outline numbers (text form: "2", "2.1", "2.1.3")
    parentOutline = parentTask.OutlineNumber
    childOutline = childTask.OutlineNumber

    ' A subtask outline starts with its parent's outline number
    If Len(childOutline) > Len(parentOutline) Then
        If Left$(childOutline, Len(parentOutline)) = parentOutline Then
            IsSubTaskOf = True
        End If
    End If
End Function

' ============================ Vue Client ============================

Private Function BuildViewClient() As String
    Dim html As String

    html = Join(Array( _
        "<div class='view-section' data-view='client'>", _
        "  <h2 style='text-align:center;margin-bottom:20px;'>DASHBOARD CLIENT - EDF RE</h2>" _
    ), vbCrLf) & vbCrLf

    ' Sélecteur (IDs distincts)
    html = html & BuildZoneSelector("zoneFilterClient", "zoneInfoClient")

    ' Section 1 : Histogramme grandes catégories
    html = html & BuildClientSection1_Histogramme() & "<div style='margin:40px 0;border-top:2px solid #ddd;'></div>" & vbCrLf
    ' Section 2 : 3 prochaines fins
    html = html & BuildClientSection2_ProchainesFins() & "<div style='margin:40px 0;border-top:2px solid #ddd;'></div>" & vbCrLf
    ' Section 3 : 3 prochains démarrages
    html = html & BuildClientSection3_ProchainsStarts() & "<div style='margin:40px 0;border-top:2px solid #ddd;'></div>" & vbCrLf
    ' Section 4 : Planning 3 semaines
    html = html & BuildClientSection4_AvancementSemaine()

    html = html & "</div>" & vbCrLf
    BuildViewClient = html
End Function

Private Function BuildClientSection1_Histogramme() As String
    Dim html As String
    Dim t As Task, child As Task
    Dim summaryTasks As Object: Set summaryTasks = CreateObject("Scripting.Dictionary")
    Dim pct As Double, recapZone As String
    Dim i As Long, j As Long, k As Variant, temp As Variant
    Dim taskName As String, taskPct As Double, taskZone As String, pctStr As String

    html = "<h3 style='font-size:14px;font-weight:bold;margin-bottom:20px;'>AVANCEMENT PAR GRANDE CATÉGORIE</h3>" & vbCrLf
    html = html & "<div class='mechanical-progress-grid' style='max-width:900px;margin:0 auto;'>" & vbCrLf

    ' Récapitulatifs niveau 2
    For Each t In ActiveProject.Tasks
        On Error Resume Next
        If Not t Is Nothing Then
            If t.Summary And t.OutlineLevel = 2 Then
                pct = 0: If Not IsEmpty(t.PercentComplete) And Not IsNull(t.PercentComplete) Then pct = CDbl(t.PercentComplete)
                ' Première feuille enfant -> zone
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

    ' Tri décroissant par %
    If summaryTasks.Count > 0 Then
        Dim keys() As Variant
        ReDim keys(summaryTasks.Count - 1)
        i = 0: For Each k In summaryTasks.keys: keys(i) = k: i = i + 1: Next k
        For i = 0 To UBound(keys) - 1
            For j = i + 1 To UBound(keys)
                If summaryTasks(keys(i))(1) < summaryTasks(keys(j))(1) Then temp = keys(i): keys(i) = keys(j): keys(j) = temp
            Next j
        Next i

        For i = 0 To UBound(keys)
            taskName = summaryTasks(keys(i))(0)
            taskPct = summaryTasks(keys(i))(1)
            taskZone = CStr(summaryTasks(keys(i))(2))
            pctStr = Replace(CStr(taskPct), ",", ".")
            html = html & "  <div class='mechanical-progress-row' data-zone='" & EncodeHTML(taskZone) & "' data-pct='" & pctStr & "'>" & vbCrLf & _
                          "    <div class='mechanical-progress-label'>" & EncodeHTML(taskName) & "</div>" & vbCrLf & _
                          "    <div class='mechanical-progress-bar-container'>" & vbCrLf & _
                          "      <div class='mechanical-progress-bar' style='width:" & pctStr & "%;'></div>" & vbCrLf & _
                          "    </div>" & vbCrLf & _
                          "    <div class='mechanical-progress-percentage'>" & Replace(CStr(Round(taskPct, 1)), ".", ",") & "%</div>" & vbCrLf & _
                          "  </div>" & vbCrLf
        Next i
    Else
        html = html & "  <p style='text-align:center;color:#666;'>Aucune tâche récapitulative trouvée</p>" & vbCrLf
    End If

    html = html & "</div>" & vbCrLf
    BuildClientSection1_Histogramme = html
End Function

Private Function BuildClientSection2_ProchainesFins() As String
    Dim html As String
    Dim t As Task, candidateTasks As Object: Set candidateTasks = CreateObject("Scripting.Dictionary")
    Dim dateToday As Date: dateToday = Date
    Dim pct As Double, sortKey As String
    Dim i As Long, k As Variant, j As Long, temp As Variant
    Dim maxRows As Long, taskName As String, taskPct As Double, taskFinish As Date, pctClass As String

    html = "<h3 style='font-size:14px;font-weight:bold;margin:40px 0 20px 0;'>PROCHAINES TÂCHES À TERMINER</h3>" & vbCrLf
    html = html & "<table class='progress-table'>" & vbCrLf
    html = html & "  <thead><tr><th>Tâche</th><th>% Avancement</th><th>Date fin prévue</th></tr></thead><tbody>" & vbCrLf

    For Each t In ActiveProject.Tasks
        On Error Resume Next
        If Not t Is Nothing Then
            If Not t.Summary Then
                pct = 0: If Not IsEmpty(t.PercentComplete) And Not IsNull(t.PercentComplete) Then pct = CDbl(t.PercentComplete)
                If pct > 0 And pct < 100 Then
                    If Not IsEmpty(t.Finish) And Not IsNull(t.Finish) Then
                        If t.Finish >= dateToday Then
                            sortKey = Format(t.Finish, "yyyymmdd") & "_" & Format(1000 - pct, "0000")
                            candidateTasks.Add sortKey, Array(t.Name, pct, t.Finish)
                        End If
                    End If
                End If
            End If
        End If
        On Error GoTo 0
    Next t

    If candidateTasks.Count > 0 Then
        Dim keys() As Variant
        ReDim keys(candidateTasks.Count - 1)
        i = 0: For Each k In candidateTasks.keys: keys(i) = k: i = i + 1: Next k
        For i = 0 To UBound(keys) - 1
            For j = i + 1 To UBound(keys)
                If keys(i) > keys(j) Then temp = keys(i): keys(i) = keys(j): keys(j) = temp
            Next j
        Next i

        maxRows = 3: If UBound(keys) + 1 < maxRows Then maxRows = UBound(keys) + 1
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
            html = html & "    <tr>" & vbCrLf & _
                          "      <td class='action-name'>" & EncodeHTML(taskName) & "</td>" & vbCrLf & _
                          "      <td class='percentage " & pctClass & "'>" & Replace(CStr(Round(taskPct, 1)), ".", ",") & "%</td>" & vbCrLf & _
                          "      <td>" & Format(taskFinish, "dd/mm/yyyy") & "</td>" & vbCrLf & _
                          "    </tr>" & vbCrLf
        Next i
    Else
        html = html & "    <tr><td colspan='3' style='text-align:center;color:#666;'>Aucune tâche en cours à terminer</td></tr>" & vbCrLf
    End If

    html = html & "  </tbody></table>" & vbCrLf
    BuildClientSection2_ProchainesFins = html
End Function

Private Function BuildClientSection3_ProchainsStarts() As String
    Dim html As String
    Dim t As Task, candidateTasks As Object: Set candidateTasks = CreateObject("Scripting.Dictionary")
    Dim dateToday As Date: dateToday = Date
    Dim pct As Double, sortKey As String, uniqueKey As String, counter As Integer
    Dim i As Long, k As Variant, j As Long, temp As Variant
    Dim maxRows As Long, taskName As String, taskStart As Date

    html = "<h3 style='font-size:14px;font-weight:bold;margin:40px 0 20px 0;'>PROCHAINES TÂCHES À DÉMARRER</h3>" & vbCrLf
    html = html & "<table class='progress-table'>" & vbCrLf
    html = html & "  <thead><tr><th>Tâche</th><th>Date démarrage prévue</th></tr></thead><tbody>" & vbCrLf

    For Each t In ActiveProject.Tasks
        On Error Resume Next
        If Not t Is Nothing Then
            If Not t.Summary Then
                pct = 0: If Not IsEmpty(t.PercentComplete) And Not IsNull(t.PercentComplete) Then pct = CDbl(t.PercentComplete)
                If pct = 0 Then
                    If Not IsEmpty(t.Start) And Not IsNull(t.Start) Then
                        If t.Start >= dateToday Then
                            sortKey = Format(t.Start, "yyyymmdd")
                            uniqueKey = sortKey: counter = 0
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

    If candidateTasks.Count > 0 Then
        Dim keys() As Variant
        ReDim keys(candidateTasks.Count - 1)
        i = 0: For Each k In candidateTasks.keys: keys(i) = k: i = i + 1: Next k
        For i = 0 To UBound(keys) - 1
            For j = i + 1 To UBound(keys)
                If keys(i) > keys(j) Then temp = keys(i): keys(i) = keys(j): keys(j) = temp
            Next j
        Next i

        maxRows = 3: If UBound(keys) + 1 < maxRows Then maxRows = UBound(keys) + 1
        For i = 0 To maxRows - 1
            taskName = candidateTasks(keys(i))(0)
            taskStart = candidateTasks(keys(i))(1)
            html = html & "    <tr>" & vbCrLf & _
                          "      <td class='action-name'>" & EncodeHTML(taskName) & "</td>" & vbCrLf & _
                          "      <td>" & Format(taskStart, "dd/mm/yyyy") & "</td>" & vbCrLf & _
                          "    </tr>" & vbCrLf
        Next i
    Else
        html = html & "    <tr><td colspan='2' style='text-align:center;color:#666;'>Toutes les tâches sont démarrées</td></tr>" & vbCrLf
    End If

    html = html & "  </tbody></table>" & vbCrLf
    BuildClientSection3_ProchainsStarts = html
End Function

Private Function BuildClientSection4_AvancementSemaine() As String
    Dim html As String
    Dim t As Task, planningData As Object: Set planningData = CreateObject("Scripting.Dictionary")
    Dim today As Date: today = Date
    Dim daysFromMonday As Integer: daysFromMonday = Weekday(today, vbMonday) - 1
    Dim week1Start As Date: week1Start = today - daysFromMonday
    Dim week4End As Date: week4End = week1Start + 27
    Dim weekNum As Integer, weekStart As Date, weekLabel As String
    Dim wNum As Integer, wStart As Date, wEndDate As Date
    Dim zone As String, entreprise As String, keyZoneEnt As String
    Dim a As Assignment, r As Resource
    Dim rhCount As Long, currentActivities As String
    Dim keyEntry As Variant, zoneData As Object, weekK As String, rhCnt As Long, activities As String
    Dim weekData As Object, zoneDict As Object

    html = "<h3 style='font-size:14px;font-weight:bold;margin:40px 0 20px 0;'>PLANNING 3 SEMAINES</h3>" & vbCrLf
    html = html & "<table class='progress-table' style='font-size:10px;width:100%;'>" & vbCrLf
    html = html & "  <thead>" & vbCrLf
    html = html & "    <tr>" & vbCrLf
    html = html & "      <th rowspan='2' style='width:80px;'>Zone</th>" & vbCrLf
    html = html & "      <th rowspan='2' style='width:100px;'>Entreprise</th>" & vbCrLf

    For weekNum = 1 To 4
        weekStart = week1Start + (weekNum - 1) * 7
        weekLabel = "S" & (5 + weekNum)
        html = html & "      <th colspan='2' style='background:#d0d0d0;font-weight:bold;'>" & weekLabel & "<br><span style='font-size:9px;font-weight:normal;'>" & Format(weekStart, "dd/mm/yyyy") & "</span></th>" & vbCrLf
    Next weekNum

    html = html & "    </tr>" & vbCrLf
    html = html & "    <tr>" & vbCrLf
    For weekNum = 1 To 4
        html = html & "      <th style='width:50px;background:#e8e8e8;'>RH</th>" & vbCrLf
        html = html & "      <th style='min-width:150px;background:#e8e8e8;'>Activité</th>" & vbCrLf
    Next weekNum
    html = html & "    </tr>" & vbCrLf
    html = html & "  </thead><tbody>" & vbCrLf

    ' Collecte par (Sous-Zone Text3 + Entreprise)
    For Each t In ActiveProject.Tasks
        On Error Resume Next
        If Not t Is Nothing And Not t.Summary Then
            If Not IsEmpty(t.Start) And Not IsNull(t.Start) And Not IsEmpty(t.Finish) And Not IsNull(t.Finish) Then
                If t.Start <= week4End And t.Finish >= week1Start Then
                    zone = IIf(Len(t.Text3) > 0, CStr(t.Text3), "Sans zone")
                    ' ?? Si votre "Entreprise" est Text5 dans vos plans, remplacez t.Text1 par t.Text5
                    entreprise = IIf(Len(t.Text1) > 0, CStr(t.Text1), "Non défini")

                    keyZoneEnt = zone & "|" & entreprise

                    If Not planningData.Exists(keyZoneEnt) Then
                        Set zoneDict = CreateObject("Scripting.Dictionary")
                        zoneDict.Add "zone", zone
                        zoneDict.Add "entreprise", entreprise
                        zoneDict.Add "weeks", CreateObject("Scripting.Dictionary")
                        planningData.Add keyZoneEnt, zoneDict
                    End If

                    For wNum = 1 To 4
                        wStart = week1Start + (wNum - 1) * 7
                        wEndDate = wStart + 6
                        If t.Start <= wEndDate And t.Finish >= wStart Then
                            Dim weekKey As String: weekKey = "S" & (5 + wNum)
                            If Not planningData(keyZoneEnt)("weeks").Exists(weekKey) Then
                                Set weekData = CreateObject("Scripting.Dictionary")
                                weekData.Add "rh", 0
                                weekData.Add "activities", ""
                                planningData(keyZoneEnt)("weeks").Add weekKey, weekData
                            End If
                            rhCount = 0
                            For Each a In t.Assignments
                                If Not a Is Nothing Then
                                    Set r = a.Resource
                                    If Not r Is Nothing Then
                                        If r.Type = pjResourceTypeWork Then rhCount = rhCount + 1
                                    End If
                                End If
                            Next a
                            If rhCount > planningData(keyZoneEnt)("weeks")(weekKey)("rh") Then
                                planningData(keyZoneEnt)("weeks")(weekKey)("rh") = rhCount
                            End If
                            currentActivities = planningData(keyZoneEnt)("weeks")(weekKey)("activities")
                            If Len(currentActivities) > 0 Then
                                If InStr(currentActivities, t.Name) = 0 Then currentActivities = currentActivities & ", " & t.Name
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

    If planningData.Count > 0 Then
        For Each keyEntry In planningData.keys
            Set zoneData = planningData(keyEntry)
            html = html & "    <tr>" & vbCrLf & _
                          "      <td class='action-name' style='text-align:center;'>" & EncodeHTML(zoneData("zone")) & "</td>" & vbCrLf & _
                          "      <td class='action-name' style='text-align:center;'>" & EncodeHTML(zoneData("entreprise")) & "</td>" & vbCrLf
            For weekNum = 1 To 4
                Dim weekKLocal As String: weekKLocal = "S" & (5 + weekNum)
                rhCnt = 0: activities = "Pas d'accès PL"
                If zoneData("weeks").Exists(weekKLocal) Then
                    rhCnt = zoneData("weeks")(weekKLocal)("rh")
                    activities = zoneData("weeks")(weekKLocal)("activities")
                    If Len(activities) = 0 Then activities = "Pas d'accès PL"
                End If
                html = html & "      <td style='text-align:center;font-weight:bold;'>" & rhCnt & "</td>" & vbCrLf & _
                              "      <td style='text-align:left;padding:5px;'>" & EncodeHTML(activities) & "</td>" & vbCrLf
            Next weekNum
            html = html & "    </tr>" & vbCrLf
        Next keyEntry
    Else
        html = html & "    <tr><td colspan='10' style='text-align:center;color:#666;'>Aucune tâche dans les 4 prochaines semaines</td></tr>" & vbCrLf
    End If

    html = html & "  </tbody></table>" & vbCrLf
    BuildClientSection4_AvancementSemaine = html
End Function

' ============================ S-Curve (MECA) ========================

Private Function BuildViewSCurve() As String
    Dim html As String
    Dim t As Task, a As Assignment, r As Resource
    Dim dateToday As Date: dateToday = Date
    Dim totalWorkPlanned As Double, totalWorkDone As Double
    Dim dateDict As Object: Set dateDict = CreateObject("Scripting.Dictionary")
    Dim metier As String, taskWork As Double, taskWorkDone As Double, taskPct As Double
    Dim qtyTotal As Double, taskDuration As Long
    Dim currentDate As Date, dateKey As String
    Dim dailyWork As Double, actualEndDate As Date, actualDuration As Long, dailyActual As Double
    Dim dayData As Object, dayDataNew As Object
    Dim plannedData As String, actualData As String
    Dim cumulPlanned As Double, cumulActual As Double
    Dim sortedDates() As String, i As Long, j As Long, temp As String, dk As Variant, dkey As String
    Dim pctPlanned As Double, pctActual As Double
    Dim debugLog As String: debugLog = "=== DEBUG S-CURVE ===" & vbCrLf

    ' Collecte MECA
    For Each t In ActiveProject.Tasks
        If Not t Is Nothing And Not t.Summary Then
            metier = IIf(Len(t.Text4) > 0, CStr(t.Text4), "")
            debugLog = debugLog & "Tache: " & t.Name & " | Text4=" & metier & " | MapGroup=" & MapGroup(metier) & vbCrLf
            If MapGroup(metier) = "meca" Then
                debugLog = debugLog & "  -> MECA detectee!" & vbCrLf
                taskWork = 0: taskWorkDone = 0: taskPct = 0
                If Not IsEmpty(t.PercentComplete) And Not IsNull(t.PercentComplete) Then taskPct = CDbl(t.PercentComplete)
                debugLog = debugLog & "  %Acheve tache=" & taskPct & vbCrLf

                On Error Resume Next
                For Each a In t.Assignments
                    If Not a Is Nothing Then
                        Set r = a.Resource
                        If Not r Is Nothing And r.Type = pjResourceTypeMaterial Then
                            qtyTotal = 0
                            If Not IsEmpty(a.Units) And Not IsNull(a.Units) Then qtyTotal = CDbl(a.Units)
                            debugLog = debugLog & "    Ressource: " & r.Name & " | Units=" & qtyTotal & vbCrLf
                            taskWork = taskWork + qtyTotal
                        End If
                    End If
                Next a
                On Error GoTo 0

                If taskWork > 0 And taskPct > 0 Then taskWorkDone = (taskWork * taskPct) / 100#
                debugLog = debugLog & "  Travail: prevu=" & taskWork & " | realise=" & taskWorkDone & vbCrLf

                If taskWork > 0 Then
                    totalWorkPlanned = totalWorkPlanned + taskWork
                    totalWorkDone = totalWorkDone + taskWorkDone

                    taskDuration = t.Finish - t.Start
                    If taskDuration <= 0 Then taskDuration = 1
                    dailyWork = taskWork / taskDuration

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

                    If taskWorkDone > 0 Then
                        If t.Finish < dateToday Then actualEndDate = t.Finish Else actualEndDate = dateToday
                        actualDuration = actualEndDate - t.Start
                        If actualDuration <= 0 Then actualDuration = 1
                        dailyActual = taskWorkDone / actualDuration

                        currentDate = t.Start
                        Do While currentDate <= actualEndDate
                            dateKey = Format(currentDate, "yyyy-mm-dd")
                            If dateDict.Exists(dateKey) Then
                                dateDict(dateKey)("actual") = dateDict(dateKey)("actual") + dailyActual
                            Else
                                Set dayDataNew = CreateObject("Scripting.Dictionary")
                                dayDataNew.Add "planned", 0#
                                dayDataNew.Add "actual", dailyActual
                                dateDict.Add dateKey, dayDataNew
                            End If
                            currentDate = currentDate + 1
                        Loop
                    End If
                End If
            End If
        End If
    Next t

    ' JSON data for Chart
    plannedData = "[": actualData = "["
    cumulPlanned = 0: cumulActual = 0
    If dateDict.Count > 0 And totalWorkPlanned > 0 Then
        ReDim sortedDates(dateDict.Count - 1)
        i = 0: For Each dk In dateDict.keys: sortedDates(i) = CStr(dk): i = i + 1: Next dk
        For i = 0 To UBound(sortedDates) - 1
            For j = i + 1 To UBound(sortedDates)
                If sortedDates(i) > sortedDates(j) Then temp = sortedDates(i): sortedDates(i) = sortedDates(j): sortedDates(j) = temp
            Next j
        Next i
        For i = 0 To UBound(sortedDates)
            dkey = sortedDates(i)
            cumulPlanned = cumulPlanned + dateDict(dkey)("planned")
            cumulActual = cumulActual + dateDict(dkey)("actual")
            pctPlanned = (cumulPlanned / totalWorkPlanned) * 100
            pctActual = (cumulActual / totalWorkPlanned) * 100
            If pctPlanned > 100 Then pctPlanned = 100
            If pctActual > 100 Then pctActual = 100
            If i > 0 Then plannedData = plannedData & ",": actualData = actualData & ","
            plannedData = plannedData & "{x:'" & dkey & "',y:" & Replace(CStr(Round(pctPlanned, 2)), ",", ".") & "}"
            actualData = actualData & "{x:'" & dkey & "',y:" & Replace(CStr(Round(pctActual, 2)), ",", ".") & "}"
        Next i
    End If
    plannedData = plannedData & "]": actualData = actualData & "]"

    ' HTML + script
    html = Join(Array( _
        "<div class='view-section' data-view='scurve'>", _
        "  <div style='padding:20px;'>", _
        "    <h2 style='text-align:center;margin-bottom:20px;'>COURBE EN S - MÉCANIQUE</h2>", _
        "    <div style='text-align:center;margin-bottom:10px;color:#666;'>", _
        "      <span style='margin-right:20px;'>?? Total prévu : " & Round(totalWorkPlanned, 0) & " unités</span>", _
        "      <span style='margin-right:20px;'>? Total réalisé : " & Round(totalWorkDone, 0) & " unités</span>", _
        "      <span>?? Avancement : " & Round((totalWorkDone / IIf(totalWorkPlanned > 0, totalWorkPlanned, 1)) * 100, 1) & "%</span>", _
        "    </div>", _
        "    <div style='max-width:1200px;margin:0 auto;height:600px;'><canvas id='scurveChart'></canvas></div>", _
        "    <div style='margin-top:20px;padding:20px;background:#f5f5f5;border:1px solid #ddd;font-family:monospace;font-size:11px;white-space:pre-wrap;'>" & EncodeHTML(debugLog) & "</div>" _
    ), vbCrLf) & vbCrLf

    html = html & "<script>" & vbCrLf
    html = html & "const plannedData = " & plannedData & ";" & vbCrLf
    html = html & "const actualData = " & actualData & ";" & vbCrLf
    html = html & Join(Array( _
        "function renderSCurve(){", _
        "  const ctx = document.getElementById('scurveChart');", _
        "  if(!ctx || window.scurveChartInstance) return;", _
        "  window.scurveChartInstance = new Chart(ctx, {", _
        "    type:'line',", _
        "    data:{datasets:[", _
        "      {label:'Prévu (Baseline)', data:plannedData, borderColor:'#4CAF50', backgroundColor:'rgba(76,175,80,0.1)', borderWidth:3, borderDash:[5,5], tension:0.4, fill:false, pointRadius:0},", _
        "      {label:'Réel (Actuel)',   data:actualData,  borderColor:'#FFD700', backgroundColor:'rgba(255,215,0,0.1)', borderWidth:3, tension:0.4, fill:false, pointRadius:3, pointHoverRadius:6}", _
        "    ]},", _
        "    options:{", _
        "      responsive:true, maintainAspectRatio:false, interaction:{mode:'index',intersect:false},", _
        "      scales:{", _
        "        x:{type:'time', time:{unit:'day'}, title:{display:true,text:'Date',font:{size:14,weight:'bold'}}, grid:{color:'rgba(0,0,0,0.1)'}},", _
        "        y:{min:0,max:100, title:{display:true,text:'% Avancement Cumulé',font:{size:14,weight:'bold'}}, grid:{color:'rgba(0,0,0,0.1)'}, ticks:{callback:(v)=>v+'%'}}", _
        "      },", _
        "      plugins:{legend:{display:true,position:'top',labels:{font:{size:13}}}, tooltip:{callbacks:{label:(ctx)=>ctx.dataset.label+': '+ctx.parsed.y.toFixed(1)+'%'}}}", _
        "    }", _
        "  });", _
        "}" _
    ), vbCrLf) & vbCrLf
    html = html & "</script>" & vbCrLf
    html = html & "  </div>" & vbCrLf & "</div>" & vbCrLf

    BuildViewSCurve = html
End Function

' =========================== HTML Footer ============================

Private Function BuildHTMLFooter() As String
    BuildHTMLFooter = "</div>" & vbCrLf & "</body>" & vbCrLf & "</html>"
End Function

' ============================== DEBUG ===============================

Private Function CreateDebugLog() As String
    Dim log As String
    Dim t As Task, zone As String, metier As String
    Dim allTasksDict As Object, zoneStatsDict As Object, globalTotal As Object
    Dim taskKey As String, taskInfo As Object
    Dim heuresPrevu As Double, heuresActuel As Double
    Dim a As Assignment, r As Resource
    Dim qCalc As Double, pctAssignment As Double
    Dim zoneTotal As Object
    Dim nameZoneDict2 As Object, taskName As Variant, tKey As Variant, tKey2 As Variant
    Dim doubleCount As Integer
    Dim ecartPrevu As Double, ecartActuel As Double
    Dim zonePct As Double
    Dim summaryCount As Integer, pctSummary As Double

    Set allTasksDict = CreateObject("Scripting.Dictionary")
    Set zoneStatsDict = CreateObject("Scripting.Dictionary")
    Set globalTotal = CreateObject("Scripting.Dictionary")
    globalTotal.Add "heuresPrevu", 0
    globalTotal.Add "heuresActuel", 0
    globalTotal.Add "taskCount", 0

    log = String(80, "=") & vbCrLf & "  ANALYSE DIAGNOSTIC COMPLETE - ZONES ET DOUBLONS" & vbCrLf & String(80, "=") & vbCrLf & vbCrLf
    log = log & "PROJET : " & ActiveProject.Name & vbCrLf
    log = log & "DATE EXPORT : " & Now() & vbCrLf & vbCrLf

    log = log & vbCrLf & String(80, "=") & vbCrLf & "ETAPE 1 : INVENTAIRE COMPLET DES TACHES" & vbCrLf & String(80, "=") & vbCrLf & vbCrLf
    For Each t In ActiveProject.Tasks
        If Not t Is Nothing And Not t.Summary Then
            zone = IIf(Len(t.Text3) > 0, CStr(t.Text3), "SANS_ZONE")
            metier = IIf(Len(t.Text4) > 0, CStr(t.Text4), "N/A")

            heuresPrevu = 0: heuresActuel = 0
            On Error Resume Next
            For Each a In t.Assignments
                If Not a Is Nothing Then
                    Set r = a.Resource
                    If Not r Is Nothing And r.Type = pjResourceTypeMaterial Then
                        qCalc = 0
                        If Not IsEmpty(a.Units) And Not IsNull(a.Units) Then qCalc = CDbl(a.Units)
                        heuresPrevu = heuresPrevu + qCalc
                    End If
                End If
            Next a
            On Error GoTo 0

            pctAssignment = 0
            If Not IsEmpty(t.PercentComplete) And Not IsNull(t.PercentComplete) Then pctAssignment = CDbl(t.PercentComplete)
            heuresActuel = heuresPrevu * pctAssignment / 100#

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

            globalTotal("heuresPrevu") = globalTotal("heuresPrevu") + heuresPrevu
            globalTotal("heuresActuel") = globalTotal("heuresActuel") + heuresActuel
            globalTotal("taskCount") = globalTotal("taskCount") + 1
        End If
    Next t

    log = log & vbCrLf & String(80, "=") & vbCrLf & "ETAPE 2 : DETECTION DES DOUBLONS (meme tache dans plusieurs zones)" & vbCrLf & String(80, "=") & vbCrLf & vbCrLf
    Set nameZoneDict2 = CreateObject("Scripting.Dictionary")
    Dim tempTaskName As String
    For Each tKey In allTasksDict.keys
        Set taskInfo = allTasksDict(tKey)
        tempTaskName = taskInfo("name")
        If Not nameZoneDict2.Exists(tempTaskName) Then Set nameZoneDict2(tempTaskName) = CreateObject("Scripting.Dictionary")
        nameZoneDict2(tempTaskName).Add tKey, taskInfo
    Next tKey
    doubleCount = 0
    For Each taskName In nameZoneDict2.keys
        If nameZoneDict2(taskName).Count > 1 Then
            doubleCount = doubleCount + 1
            log = log & "DOUBLON #" & doubleCount & " : " & taskName & vbCrLf & "  Present dans " & nameZoneDict2(taskName).Count & " zones :" & vbCrLf
            For Each tKey2 In nameZoneDict2(taskName).keys
                Set taskInfo = nameZoneDict2(taskName)(tKey2)
                log = log & "    - Zone: " & taskInfo("zone") & " | Heures: " & Format(taskInfo("heuresPrevu"), "0.00") & "h" & vbCrLf
            Next tKey2
            log = log & vbCrLf
        End If
    Next taskName
    If doubleCount = 0 Then log = log & "OK - Aucun doublon detecte" & vbCrLf Else log = log & "ALERTE - TOTAL : " & doubleCount & " taches en doublon" & vbCrLf

    log = log & vbCrLf & String(80, "=") & vbCrLf & "ETAPE 3 : STATISTIQUES PAR ZONE (filtrage individuel)" & vbCrLf & String(80, "=") & vbCrLf & vbCrLf
    Dim zKey As Variant
    For Each zKey In zoneStatsDict.keys
        Set zoneTotal = zoneStatsDict(zKey)
        log = log & "+--- ZONE : " & zKey & " " & String(60 - Len(CStr(zKey)), "-") & "+" & vbCrLf & _
                    "| Nombre de taches : " & zoneTotal("taskCount") & vbCrLf & _
                    "| Heures prevues   : " & Format(zoneTotal("heuresPrevu"), "#,##0.00") & " h" & vbCrLf & _
                    "| Heures actuelles : " & Format(zoneTotal("heuresActuel"), "#,##0.00") & " h" & vbCrLf
        zonePct = 0
        If zoneTotal("heuresPrevu") > 0 Then zonePct = (zoneTotal("heuresActuel") / zoneTotal("heuresPrevu")) * 100
        log = log & "| Avancement       : " & Format(zonePct, "0.00") & " %" & vbCrLf & "+" & String(70, "-") & "+" & vbCrLf & vbCrLf
    Next zKey

    log = log & vbCrLf & String(80, "=") & vbCrLf & "ETAPE 4 : COMPARAISON DES MODES DE FILTRAGE" & vbCrLf & String(80, "=") & vbCrLf & vbCrLf
    Dim sumZonesPrevu As Double, sumZonesActuel As Double
    For Each zKey In zoneStatsDict.keys
        Set zoneTotal = zoneStatsDict(zKey)
        sumZonesPrevu = sumZonesPrevu + zoneTotal("heuresPrevu")
        sumZonesActuel = sumZonesActuel + zoneTotal("heuresActuel")
    Next zKey
    log = log & "MODE 'TOUTES AGREGEES' (toutes taches uniques) :" & vbCrLf & _
                "  Taches    : " & globalTotal("taskCount") & vbCrLf & _
                "  Heures P. : " & Format(globalTotal("heuresPrevu"), "#,##0.00") & " h" & vbCrLf & _
                "  Heures A. : " & Format(globalTotal("heuresActuel"), "#,##0.00") & " h" & vbCrLf & vbCrLf & _
                "SOMME DES ZONES (addition des filtres individuels) :" & vbCrLf & _
                "  Heures P. : " & Format(sumZonesPrevu, "#,##0.00") & " h" & vbCrLf & _
                "  Heures A. : " & Format(sumZonesActuel, "#,##0.00") & " h" & vbCrLf & vbCrLf

    ecartPrevu = sumZonesPrevu - globalTotal("heuresPrevu")
    ecartActuel = sumZonesActuel - globalTotal("heuresActuel")
    log = log & "ECART (Somme zones - Toutes agregees) :" & vbCrLf & _
                "  Heures P. : " & Format(ecartPrevu, "#,##0.00") & " h" & IIf(ecartPrevu = 0, " OK", IIf(ecartPrevu > 0, " ALERTE SURPLUS", " ALERTE MANQUE")) & vbCrLf & _
                "  Heures A. : " & Format(ecartActuel, "#,##0.00") & " h" & IIf(ecartActuel = 0, " OK", IIf(ecartActuel > 0, " ALERTE SURPLUS", " ALERTE MANQUE")) & vbCrLf & vbCrLf

    log = log & vbCrLf & String(80, "=") & vbCrLf & "ETAPE 5 : TACHES RECAPITULATIVES" & vbCrLf & String(80, "=") & vbCrLf & vbCrLf
    Dim tSum As Task
    summaryCount = 0
    For Each tSum In ActiveProject.Tasks
        If Not tSum Is Nothing And tSum.Summary Then
            summaryCount = summaryCount + 1
            pctSummary = 0: If Not IsEmpty(tSum.PercentComplete) And Not IsNull(tSum.PercentComplete) Then pctSummary = CDbl(tSum.PercentComplete)
            metier = IIf(Len(tSum.Text4) > 0, CStr(tSum.Text4), "N/A")
            zone = IIf(Len(tSum.Text3) > 0, CStr(tSum.Text3), "N/A")
            log = log & summaryCount & ". " & tSum.Name & vbCrLf & _
                        "   Zone: " & zone & " | Metier: " & metier & " | OutlineLevel: " & tSum.OutlineLevel & " | %: " & pctSummary & "%" & vbCrLf & vbCrLf
        End If
    Next tSum

    log = log & vbCrLf & String(80, "=") & vbCrLf & "ETAPE 6 : COHERENCE VUE METIERS vs DASHBOARD CLIENT" & vbCrLf & String(80, "=") & vbCrLf & vbCrLf & _
                "  >> Text3 = Zone | Text4 = Metier" & vbCrLf & vbCrLf
    log = log & LogMetierSection("vrd", "TRANCHEES  (VRD)")
    log = log & LogMetierSection("meca", "STRUCTURES (MECA)")
    log = log & LogMetierSection("elec", "ELECTRICITE (ELEC)")

    log = log & vbCrLf & String(80, "=") & vbCrLf & "DIAGNOSTIC FINAL" & vbCrLf & String(80, "=") & vbCrLf & vbCrLf
    If ecartPrevu > 0 Or ecartActuel > 0 Then
        log = log & "ALERTE PROBLEME DETECTE : Les doublons sont comptes plusieurs fois dans les filtres de zones." & vbCrLf & _
                    "   Solution recommandee : Utiliser le mode 'Toutes agregees' pour une vision globale correcte." & vbCrLf
    Else
        log = log & "OK COHERENCE VALIDEE : Pas de doublons, les filtres zones donnent des resultats coherents." & vbCrLf
    End If
    log = log & vbCrLf & "Fin du diagnostic." & vbCrLf

    CreateDebugLog = log
End Function

Private Function LogMetierSection(ByVal metierType As String, ByVal metierLabel As String) As String
    Dim log As String
    Dim t As Task, sub1 As Task
    Dim hasMatchingLot As Boolean
    Dim taskZone As String, subLot As String
    Dim pct As Double
    Dim foundList As Object: Set foundList = CreateObject("Scripting.Dictionary")
    Dim key As Variant
    Dim tName As String, tPct As Double, tZone As String
    Dim tLevel As Integer, tOutline As String
    Dim subCountMatch As Long, subCountOther As Long, subCountNone As Long
    Dim sumPctMatch As Double, pctPur As Double
    Dim otherMetiers As String, subPct As Double
    Dim isMatch As Boolean, totalSub As Long
    Dim alertMsg As String

    log = "+--- " & metierLabel & " " & String(74 - Len(metierLabel), "-") & vbCrLf

    For Each t In ActiveProject.Tasks
        If Not t Is Nothing And t.Summary And t.OutlineLevel > 1 Then
            hasMatchingLot = False
            taskZone = "SANS_ZONE"

            For Each sub1 In ActiveProject.Tasks
                If Not sub1 Is Nothing And Not sub1.Summary Then
                    If sub1.OutlineLevel > t.OutlineLevel And sub1.ID > t.ID And _
                       (sub1.OutlineParent Is Nothing Or sub1.OutlineParent.ID = t.ID Or IsSubTaskOf(sub1, t)) Then
                        subLot = LCase$(Trim$(IIf(Len(sub1.Text4) > 0, CStr(sub1.Text4), "")))
                        If (metierType = "vrd" And InStr(subLot, "vrd") > 0) Or _
                           (metierType = "meca" And (InStr(subLot, "meca") > 0 Or InStr(subLot, "mecanique") > 0)) Or _
                           (metierType = "elec" And (InStr(subLot, "elec") > 0 Or InStr(subLot, "electrique") > 0)) Then
                            hasMatchingLot = True
                            If taskZone = "SANS_ZONE" And Len(sub1.Text3) > 0 Then taskZone = CStr(sub1.Text3)
                        End If
                    End If
                End If
            Next sub1

            If hasMatchingLot And Not foundList.Exists(t.ID) Then
                pct = 0: If Not IsEmpty(t.PercentComplete) And Not IsNull(t.PercentComplete) Then pct = CDbl(t.PercentComplete)
                foundList.Add t.ID, Array(t.Name, pct, taskZone, t.OutlineLevel, t.OutlineNumber)
            End If
        End If
    Next t

    log = log & "| Recap detectees pour ce metier : " & foundList.Count & vbCrLf
    If foundList.Count = 0 Then
        log = log & "| => AUCUNE - verifier que Text4 des sous-taches contient '" & metierType & "'" & vbCrLf & "+" & String(79, "-") & vbCrLf & vbCrLf
        LogMetierSection = log
        Exit Function
    End If
    log = log & "|" & vbCrLf

    For Each key In foundList.keys
        tName = foundList(key)(0)
        tPct = foundList(key)(1)
        tZone = foundList(key)(2)
        tLevel = foundList(key)(3)
        tOutline = foundList(key)(4)

        log = log & "| [ID=" & key & " Niv=" & tLevel & " Plan=" & tOutline & "] " & tName & vbCrLf & _
                    "|   Zone : " & tZone & vbCrLf & _
                    "|   % AFFICHE dans Vue Metiers  = " & Format(tPct, "0.0") & "% (PercentComplete GLOBAL MSProject - tous metiers confondus)" & vbCrLf

        subCountMatch = 0: subCountOther = 0: subCountNone = 0
        sumPctMatch = 0: otherMetiers = ""

        For Each sub1 In ActiveProject.Tasks
            If Not sub1 Is Nothing And Not sub1.Summary Then
                If IsSubTaskOf(sub1, ActiveProject.Tasks(CLng(key))) Then
                    subPct = 0: If Not IsEmpty(sub1.PercentComplete) And Not IsNull(sub1.PercentComplete) Then subPct = CDbl(sub1.PercentComplete)
                    subLot = LCase$(Trim$(IIf(Len(sub1.Text4) > 0, CStr(sub1.Text4), "")))
                    isMatch = False
                    If (metierType = "vrd" And InStr(subLot, "vrd") > 0) Then isMatch = True
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
        log = log & " | " & subCountMatch & " " & UCase$(metierType)
        If subCountOther > 0 Then log = log & " | " & subCountOther & " AUTRE (" & otherMetiers & ")"
        If subCountNone > 0 Then log = log & " | " & subCountNone & " sans Text4"
        log = log & vbCrLf

        alertMsg = ""
        If subCountOther > 0 Then
            pctPur = IIf(subCountMatch > 0, sumPctMatch / subCountMatch, 0)
            log = log & "|   % pur " & UCase$(metierType) & " (sous-taches " & metierType & " seulement) = " & Format(pctPur, "0.0") & "%" & vbCrLf
            alertMsg = "INCOHERENCE : recap mixte (" & otherMetiers & "). Affiche=" & Format(tPct, "0.0") & "% vs Pur-" & UCase$(metierType) & "=" & Format(pctPur, "0.0") & "%."
            alertMsg = alertMsg & " Ecart=" & Format(Abs(tPct - pctPur), "0.0") & "%."
        ElseIf totalSub = 0 Then
            alertMsg = "ATTENTION : aucune sous-tache directe trouvee (verifier niveaux de plan OutlineNumber)."
        End If

        If Len(alertMsg) > 0 Then
            log = log & "|   >>> " & alertMsg & vbCrLf
        Else
            log = log & "|   >>> OK : toutes les sous-taches sont du metier " & UCase$(metierType) & ", % coherent." & vbCrLf
        End If
        log = log & "|" & vbCrLf
    Next key

    log = log & "+" & String(79, "-") & vbCrLf & vbCrLf
    LogMetierSection = log
End Function

' ============================== Utils ===============================

Private Function MapGroup(ByVal s As String) As String
    Dim k As String: k = LCase$(Trim$(s))
    If InStr(k, "elec") > 0 Or InStr(k, "electrique") > 0 Then
        MapGroup = "elec"
    ElseIf InStr(k, "mec") > 0 Or InStr(k, "mecanique") > 0 Then
        MapGroup = "meca"
    Else
        MapGroup = ""
    End If
End Function

Private Function EncodeHTML(ByVal rawText As String) As String
    Dim o As String
    o = rawText
    o = Replace(o, "&", "&amp;")
    o = Replace(o, "<", "&lt;")
    o = Replace(o, ">", "&gt;")
    o = Replace(o, """", "&quot;")
    o = Replace(o, "'", "&#39;")
    EncodeHTML = o
End Function
