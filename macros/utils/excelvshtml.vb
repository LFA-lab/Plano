Option Explicit

' ===============================================
' Export HTML "moderne" (style 1600-like) + logo
' ===============================================
Public Sub ExportFeuille_HTML_1600()
    Dim ws As Worksheet, rng As Range
    Dim outPath As String, fileName As String
    Dim lastRow As Long, lastCol As Long
    Dim r As Long, c As Long, i As Long
    Dim sb() As String, valText As String
    Set ws = ActiveSheet

    ' Plage exportée : sélection si présente, sinon UsedRange
    If TypeName(Selection) = "Range" Then
        Set rng = Selection
    Else
        Set rng = ws.UsedRange
    End If
    If rng Is Nothing Then
        MsgBox "Aucune donnée à exporter.", vbExclamation
        Exit Sub
    End If

    fileName = CleanFileName(ws.Name) & "_" & Format(Now, "yyyymmdd_HHMM") & ".html"
    outPath = Environ$("USERPROFILE") & "\Downloads\" & fileName

    lastRow = rng.Rows.Count
    lastCol = rng.Columns.Count

    ReDim sb(0 To 0): i = -1
    AddLine sb, i, "<!DOCTYPE html>"
    AddLine sb, i, "<html lang=""fr""><head>"
    AddLine sb, i, "  <meta charset=""utf-8"">"
    AddLine sb, i, "  <meta name=""viewport"" content=""width=device-width, initial-scale=1"">"
    AddLine sb, i, "  <title>" & HtmlEncode(ws.Parent.Name & " — " & ws.Name) & "</title>"
    AddLine sb, i, "  <style>"
    AddLine sb, i, "    :root{"
    AddLine sb, i, "      --bg:#0b0e14; --surface:#101521; --surface-2:#131a28;"
    AddLine sb, i, "      --fg:#e7eaf0; --muted:#9aa3b2; --border:#283248;"
    AddLine sb, i, "      --accent:#00A3E0; /* bleu Omexom-like */"
    AddLine sb, i, "      --radius:14px;"
    AddLine sb, i, "    }"
    AddLine sb, i, "    html,body{height:100%}"
    AddLine sb, i, "    body{margin:0;font:15px/1.55 system-ui,Segoe UI,Roboto,Arial,sans-serif;color:var(--fg);"
    AddLine sb, i, "         background: radial-gradient(1200px 600px at 10% -10%, rgba(0,163,224,.07), transparent 50%),"
    AddLine sb, i, "                     radial-gradient(1000px 500px at 110% 10%, rgba(124,124,255,.06), transparent 45%),"
    AddLine sb, i, "                     var(--bg);}"
    AddLine sb, i, "    .page{max-width:1200px;margin:40px auto;padding:0 20px}"
    AddLine sb, i, "    .hero{position:sticky;top:0;z-index:10;backdrop-filter:blur(10px);"
    AddLine sb, i, "          background:color-mix(in oklab, var(--surface), transparent 35%);"
    AddLine sb, i, "          border:1px solid var(--border);border-radius:16px;padding:14px 16px;margin-bottom:18px;"
    AddLine sb, i, "          box-shadow:0 8px 24px rgba(0,0,0,.35);}"
    AddLine sb, i, "    .hero-row{display:flex;align-items:center;gap:14px;}"
    AddLine sb, i, ""
    AddLine sb, i, "    .title{font-size:18px;font-weight:700;letter-spacing:.2px}"
    AddLine sb, i, "    .subtitle{font-size:12px;color:var(--muted)}"
    AddLine sb, i, "    .card{background:color-mix(in oklab, var(--surface), transparent 12%);"
    AddLine sb, i, "          border:1px solid var(--border);border-radius:var(--radius);padding:14px;"
    AddLine sb, i, "          box-shadow:0 10px 30px rgba(0,0,0,.25)}"
    AddLine sb, i, "    .table-wrap{overflow:auto;border-radius:10px;border:1px solid var(--border)}"
    AddLine sb, i, "    table{width:100%;border-collapse:separate;border-spacing:0;min-width:720px}"
    AddLine sb, i, "    thead th{position:sticky;top:0;background:linear-gradient(180deg, #151b28, #121826);"
    AddLine sb, i, "             color:#eaf2fb;font-weight:700;text-align:left;padding:10px 12px;"
    AddLine sb, i, "             border-bottom:1px solid var(--border);cursor:pointer;user-select:none}"
    AddLine sb, i, "    th:first-child{border-top-left-radius:10px}"
    AddLine sb, i, "    th:last-child{border-top-right-radius:10px}"
    AddLine sb, i, "    tbody td{padding:9px 12px;border-bottom:1px solid #20283a;vertical-align:top}"
    AddLine sb, i, "    tbody tr:hover{background:rgba(0,163,224,.06)}"
    AddLine sb, i, "    .num{text-align:right;white-space:nowrap;}"
    AddLine sb, i, "    .chip{display:inline-block;border:1px solid color-mix(in oklab, var(--accent), white 70%);"
    AddLine sb, i, "          color:color-mix(in oklab, var(--accent), white 10%);padding:2px 8px;border-radius:999px;font-size:12px}"
    AddLine sb, i, "    a{color:var(--accent);text-decoration:none} a:hover{text-decoration:underline}"
    AddLine sb, i, "    .meta{color:var(--muted)}"
    AddLine sb, i, "    .sort-ind{opacity:.7;margin-left:6px}"
    AddLine sb, i, "    @media (max-width:768px){.title{font-size:16px}}"
    AddLine sb, i, "  </style>"
    AddLine sb, i, "</head><body>"
    AddLine sb, i, "  <div class=""page"">"

    ' --- Header / Hero ---
    AddLine sb, i, "    <div class=""hero""><div class=""hero-row"">"
    AddLine sb, i, "      <div class=""chip"">OMEXOM</div>"
    AddLine sb, i, "      <div class=""head"">"
    AddLine sb, i, "        <div class=""title"">Export "" & HtmlEncode(ws.Name) & ""</div>"
    AddLine sb, i, "        <div class=""subtitle meta"">Classeur: " & HtmlEncode(ws.Parent.Name) & " • Plage: " & HtmlEncode(rng.Address(False, False)) & " • Généré: " & HtmlEncode(Format(Now, "yyyy-mm-dd HH:MM")) & "</div>"
    AddLine sb, i, "      </div>"
    AddLine sb, i, "    </div></div>"

    ' --- Table ---
    AddLine sb, i, "    <div class=""card"">"
    AddLine sb, i, "      <div class=""table-wrap"">"
    AddLine sb, i, "        <table id=""grid"">"
    AddLine sb, i, "          <thead><tr>"
    Dim colHeader As String
    Dim cIx As Long
    For c = 1 To lastCol
        colHeader = GetHeaderText(rng, c)
        AddLine sb, i, "            <th data-col=""" & c & """>" & HtmlEncode(colHeader) & "<span class=""sort-ind"">↕</span></th>"
    Next c
    AddLine sb, i, "          </tr></thead>"
    AddLine sb, i, "          <tbody>"
    Dim cell As Range
    For r = 1 To lastRow
        AddLine sb, i, "            <tr>"
        For c = 1 To lastCol
            Set cell = rng.Cells(r, c)
            valText = CStr(cell.Text)
            If IsNumeric(cell.Value) And Len(valText) > 0 Then
                AddLine sb, i, "              <td class=""num"">" & EncodeWithLinks(HtmlEncode(valText)) & "</td>"
            Else
                AddLine sb, i, "              <td>" & EncodeWithLinks(HtmlEncode(valText)) & "</td>"
            End If
        Next c
        AddLine sb, i, "            </tr>"
    Next r
    AddLine sb, i, "          </tbody>"
    AddLine sb, i, "        </table>"
    AddLine sb, i, "      </div>"
    AddLine sb, i, "    </div>"

    ' --- Tri des colonnes (vanilla JS) ---
    AddLine sb, i, "    <script>"
    AddLine sb, i, "      (function(){"
    AddLine sb, i, "        const table=document.getElementById('grid');"
    AddLine sb, i, "        if(!table) return;"
    AddLine sb, i, "        const getCell=(row,idx)=>row.children[idx-1].innerText.trim();"
    AddLine sb, i, "        const asNum=s=>{const n=Number(String(s).replace(/\s/g,'')); return isNaN(n)?null:n;};"
    AddLine sb, i, "        const cmp=(a,b)=> (a>b)-(a<b);"
    AddLine sb, i, "        table.querySelectorAll('thead th').forEach((th,ix)=>{"
    AddLine sb, i, "          let dir=1;"
    AddLine sb, i, "          th.addEventListener('click',()=>{"
    AddLine sb, i, "            const col=ix+1;"
    AddLine sb, i, "            const rows=[...table.tBodies[0].rows];"
    AddLine sb, i, "            rows.sort((ra,rb)=>{"
    AddLine sb, i, "              const sa=getCell(ra,col), sb=getCell(rb,col);"
    AddLine sb, i, "              const na=asNum(sa), nb=asNum(sb);"
    AddLine sb, i, "              return (na!=null && nb!=null ? (na-nb) : cmp(sa.toLowerCase(), sb.toLowerCase()))*dir;"
    AddLine sb, i, "            });"
    AddLine sb, i, "            dir*=-1;"
    AddLine sb, i, "            rows.forEach(r=>table.tBodies[0].appendChild(r));"
    AddLine sb, i, "          });"
    AddLine sb, i, "        });"
    AddLine sb, i, "      })();"
    AddLine sb, i, "    </script>"

    AddLine sb, i, "  </div>"
    AddLine sb, i, "</body></html>"

    On Error GoTo FailWrite
    WriteUtf8 outPath, Join(sb, vbCrLf)

    MsgBox "Export HTML (design moderne) terminé :" & vbCrLf & outPath, vbInformation
    Exit Sub

FailWrite:
    MsgBox "Erreur écriture fichier : " & Err.Description, vbCritical
End Sub


' ==================
' Helpers / Utiles
' ==================

Private Function CleanFileName(ByVal s As String) As String
    Dim bad As Variant, ch As Variant
    bad = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    For Each ch In bad
        s = Replace$(s, CStr(ch), "_")
    Next
    CleanFileName = s
End Function

Private Sub AddLine(ByRef arr() As String, ByRef idx As Long, ByVal line As String)
    idx = idx + 1
    ReDim Preserve arr(0 To idx)
    arr(idx) = line
End Sub

Private Function HtmlEncode(ByVal s As String) As String
    s = Replace$(s, "&", "&amp;")
    s = Replace$(s, "<", "&lt;")
    s = Replace$(s, ">", "&gt;")
    s = Replace$(s, """", "&quot;")
    HtmlEncode = s
End Function

Private Function EncodeWithLinks(ByVal s As String) As String
    Dim re As Object
    Set re = CreateObject("VBScript.Regexp")
    re.Pattern = "((https?://|www\.)[^\s<]+)"
    re.Global = True
    re.IgnoreCase = True
    EncodeWithLinks = re.Replace(s, "<a href=""$1"">$1</a>")
    EncodeWithLinks = Replace$(EncodeWithLinks, "href=""www.", "href=""https://www.")
End Function

Private Function GetHeaderText(rng As Range, ByVal colIndex As Long) As String
    Dim firstRow As Range
    Set firstRow = rng.Rows(1)
    Dim t As String
    t = CStr(firstRow.Cells(1, colIndex).Text)
    If Len(t) > 0 Then
        GetHeaderText = t
    Else
        GetHeaderText = Split(Cells(1, colIndex).Address(False, False), "1")(0)
    End If
End Function

Private Sub WriteUtf8(ByVal filePath As String, ByVal content As String)
    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    With stm
        .Type = 2: .Charset = "utf-8": .Open
        .WriteText content
        .SaveToFile filePath, 2 ' adSaveCreateOverWrite
        .Close
    End With
End Sub
