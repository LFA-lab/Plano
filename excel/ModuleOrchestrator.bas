Attribute VB_Name = "ModuleOrchestrator"
Option Explicit
' =====================================================================
' ModuleOrchestrator.bas  (Branch 6 – clean single module)
' - Orchestrator using previous capabilities (UI + downloader + MS Project)
' - Step management with progress messages and button disable/enable
' - Reads config from JSON (tiny built-in parser; no external libs)
' - Comprehensive Debug.Print logging
' =====================================================================

' ============================= CONFIG =================================
' Where to read JSON config from (relative to workbook folder)
Private Const CFG_REL_PATH As String = "config\orchestrator.json"

' Default UI (overridden by JSON if present)
Private Const UI_SHEET_NAME_DEFAULT As String = "Feuil1"
Private Const UI_STATUS_CELL_DEFAULT As String = "B5"

' Downloads target base folder
'   Windows: %USERPROFILE%\Downloads\omexom
'   Linux  : ~/Downloads/omexom
' =====================================================================

' =========================== PUBLIC MACROS =============================

' Create a default config file (idempotent). Run once if you don't have config yet.
Public Sub Orchestrator_WriteDefaultConfig()
    Dim cfgPath As String, folderPath As String
    cfgPath = WorkbookPathJoin(CFG_REL_PATH)
    folderPath = ParentPath(cfgPath)
    MkDirs folderPath

    Dim json As String
    json = _
    "{" & vbCrLf & _
    "  ""ui"": {""sheet"": ""Feuil1"", ""status_cell"": ""B5""}," & vbCrLf & _
    "  ""downloads"": [" & vbCrLf & _
    "    { ""url"": ""https://raw.githubusercontent.com/lfa-lab/Omexom/main/TemplateProject_v1.mpt"", ""save_as"": ""TemplateProject_v1.mpt"" }," & vbCrLf & _
    "    { ""url"": ""https://raw.githubusercontent.com/lfa-lab/Omexom/main/macros/Macro%20MSP/FichierBaseArriv%C3%A9e.mpp"", ""save_as"": ""FichierBaseArrivée.mpp"" }" & vbCrLf & _
    "  ]," & vbCrLf & _
    "  ""msproject"": { ""open_template"": ""TemplateProject_v1.mpt"", ""run_macro"": false, ""macro_name"": ""TemplateProject_v1!Module1.SampleMacro"" }" & vbCrLf & _
    "}" & vbCrLf

    WriteAllText cfgPath, json
    Debug.Print TS(), "[CFG] Wrote default config -> " & cfgPath
    MsgBox "Default config written to:" & vbCrLf & cfgPath, vbInformation, "Orchestrator"
End Sub

' Validate reading + parsing config
Public Sub Orchestrator_TestConfig()
    Dim cfg As String: cfg = LoadConfigJson()
    If Len(cfg) = 0 Then
        Debug.Print TS(), "[CFG][ERROR] Config not found."
        MsgBox "Config not found. Run Orchestrator_WriteDefaultConfig first.", vbExclamation
        Exit Sub
    End If

    Dim uiSheet As String, uiCell As String
    uiSheet = JsonGetString(cfg, "ui.sheet", UI_SHEET_NAME_DEFAULT)
    uiCell = JsonGetString(cfg, "ui.status_cell", UI_STATUS_CELL_DEFAULT)
    Debug.Print TS(), "[CFG] ui.sheet=", uiSheet, "  ui.status_cell=", uiCell

    Dim urls() As String, saves() As String, n As Long
    n = JsonGetDownloads(cfg, urls, saves)
    Debug.Print TS(), "[CFG] downloads count=", n
    Dim i As Long
    For i = 0 To n - 1
        Debug.Print TS(), "   - url=", urls(i), "  save_as=", saves(i)
    Next

    Dim tpl As String, runMacro As Boolean, mName As String
    tpl = JsonGetString(cfg, "msproject.open_template", "TemplateProject_v1.mpt")
    runMacro = JsonGetBool(cfg, "msproject.run_macro", False)
    mName = JsonGetString(cfg, "msproject.macro_name", "")
    Debug.Print TS(), "[CFG] msproject.open_template=", tpl, "  run_macro=", runMacro, "  macro_name=", mName

    MsgBox "Config parsed OK. See Immediate Window (Ctrl+G).", vbInformation
End Sub

' Main orchestrator: runs all steps with progress + logging
Public Sub Orchestrator_Run()
    Dim cfg As String: cfg = LoadConfigJson()
    If Len(cfg) = 0 Then
        MsgBox "Config not found. Run Orchestrator_WriteDefaultConfig first.", vbExclamation
        Exit Sub
    End If

    ' === resolve UI targets from config (with defaults) ===
    Dim uiSheetName As String, uiCell As String
    uiSheetName = JsonGetString(cfg, "ui.sheet", UI_SHEET_NAME_DEFAULT)
    uiCell = JsonGetString(cfg, "ui.status_cell", UI_STATUS_CELL_DEFAULT)

    Dim ws As Worksheet: Set ws = SheetOrNothing(uiSheetName)
    If ws Is Nothing Then MsgBox "UI sheet '" & uiSheetName & "' not found.", vbExclamation: Exit Sub

    Dim disabled As Boolean: disabled = False
    On Error GoTo EH

    Debug.Print TS(), "=== ORCHESTRATION START ==="
    UI_Status ws, uiCell, "? Processing..."

    ' Disable UI buttons during run
    disabled = DisableButtons(ws, True)

    ' === Step 1: downloads ===
    Debug.Print TS(), "[STEP 1] DOWNLOADS"
    UI_Status ws, uiCell, "Step 1/3: downloading files..."
    Dim urls() As String, saves() As String, n As Long, i As Long
    n = JsonGetDownloads(cfg, urls, saves)
    If n = 0 Then Debug.Print TS(), "[STEP 1][WARN] No downloads listed in config."
    For i = 0 To n - 1
        Debug.Print TS(), "  -> GET ", urls(i), "  as  ", saves(i)
        UI_Status ws, uiCell, "Downloading: " & saves(i)
        If Not HttpGetToFile(urls(i), BuildOutputPath(saves(i))) Then
            Debug.Print TS(), "[STEP 1][ERROR] Download failed: ", urls(i)
            UI_Status ws, uiCell, "? Failed: " & saves(i)
            GoTo FinallyLabel
        Else
            Debug.Print TS(), "[STEP 1][OK] -> ", BuildOutputPath(saves(i))
        End If
    Next i

    ' === Step 2: MS Project open / optional macro ===
    Debug.Print TS(), "[STEP 2] MS PROJECT"
    UI_Status ws, uiCell, "Step 2/3: opening MS Project..."

    Dim tpl As String, runMacro As Boolean, mName As String
    tpl = JsonGetString(cfg, "msproject.open_template", "TemplateProject_v1.mpt")
    runMacro = JsonGetBool(cfg, "msproject.run_macro", False)
    mName = JsonGetString(cfg, "msproject.macro_name", "")

    Dim localTpl As String: localTpl = ProjectTemplatePath(tpl)
    Debug.Print TS(), "[STEP 2] template resolved: ", localTpl
    MSP_Open localTpl, mName, (runMacro And Len(Trim$(mName)) > 0)

    ' === Step 3: finalize ===
    Debug.Print TS(), "[STEP 3] FINALIZE"
    UI_Status ws, uiCell, "Step 3/3: finalizing..."
    ' (Put any final exports or cleanup here)
    DoEvents

    UI_Status ws, uiCell, "? Complete!"
    Debug.Print TS(), "=== ORCHESTRATION END ==="
    GoTo FinallyLabel

EH:
    Debug.Print TS(), "[ERROR] ", Err.Number, " - ", Err.Description
    UI_Status ws, uiCell, "? Error: " & Err.Description
FinallyLabel:
    If disabled Then DisableButtons ws, False
End Sub

' Quick connectivity test to GitHub RAW
Public Sub Orchestrator_TestDownloadPing()
    Dim url As String, outPath As String
    url = "https://raw.githubusercontent.com/lfa-lab/Omexom/main/README.md"
    outPath = BuildOutputPath("README.md")
    Debug.Print TS(), "[PING] ", url
    If HttpGetToFile(url, outPath) Then
        Debug.Print TS(), "[PING][OK] -> ", outPath
        MsgBox "Download OK ? " & outPath, vbInformation
    Else
        Debug.Print TS(), "[PING][FAIL]"
        MsgBox "Download failed. See Immediate Window (Ctrl+G).", vbExclamation
    End If
End Sub


' ====================== UI (status + buttons) ==========================

Private Sub UI_Status(ByVal ws As Worksheet, ByVal cellAddr As String, ByVal txt As String)
    ws.Range(cellAddr).Value = txt
    DoEvents
    Debug.Print TS(), "[UI] ", txt
End Sub

' Disable/enable Form buttons and ActiveX CommandButtons on the sheet
Private Function DisableButtons(ByVal ws As Worksheet, ByVal stateDisable As Boolean) As Boolean
    On Error Resume Next
    Dim hadAny As Boolean: hadAny = False
    Dim shp As Shape, ole As OLEObject

    For Each shp In ws.Shapes
        If shp.Type = msoFormControl Then
            If shp.FormControlType = xlButtonControl Then
                shp.ControlFormat.Enabled = Not stateDisable
                hadAny = True
            End If
        End If
    Next shp

    For Each ole In ws.OLEObjects
        If Not ole.Object Is Nothing Then
            If TypeName(ole.Object) = "CommandButton" Then
                ole.Object.Enabled = Not stateDisable
                hadAny = True
            End If
        End If
    Next ole

    DisableButtons = hadAny
End Function

Private Function SheetOrNothing(ByVal name As String) As Worksheet
    On Error Resume Next
    Set SheetOrNothing = ThisWorkbook.Worksheets(name)
End Function


' ========================= MS PROJECT (COM) ============================

' Opens template; optionally runs macro; closes (no save).
Private Sub MSP_Open(ByVal templatePath As String, _
                     Optional ByVal macroName As String = "", _
                     Optional ByVal runMacro As Boolean = False)

    If Not IsWindowsOS Then
        Debug.Print TS(), "[MSP] Non-Windows OS. COM requires Windows + MS Project."
        Exit Sub
    End If

    On Error GoTo EH
    Dim app As Object, proj As Object

    Debug.Print TS(), "[MSP] Starting MS Project..."
    Set app = CreateObject("MSProject.Application")
    app.Visible = True

    Debug.Print TS(), "[MSP] Opening: ", templatePath
    app.FileOpen name:=templatePath

    Set proj = app.ActiveProject
    Debug.Print TS(), "[MSP] ActiveProject: ", proj.name

    If runMacro And Len(Trim$(macroName)) > 0 Then
        Debug.Print TS(), "[MSP] Running macro: ", macroName
        app.Run macroName
        Debug.Print TS(), "[MSP] Macro finished."
    Else
        Debug.Print TS(), "[MSP] No macro requested (smoke test)."
    End If

    Debug.Print TS(), "[MSP] Closing (no save)."
    app.FileClose pjDoNotSave

    Debug.Print TS(), "[MSP] Quitting."
    app.Quit

    Set proj = Nothing
    Set app = Nothing
    Debug.Print TS(), "[MSP] DONE."
    Exit Sub
EH:
    Debug.Print TS(), "[MSP][ERROR] ", Err.Number, " - ", Err.Description
    On Error Resume Next
    If Not app Is Nothing Then
        app.FileClose pjDoNotSave
        app.Quit
    End If
End Sub

' Finds template near workbook first, then \templates
Private Function ProjectTemplatePath(ByVal fileName As String) As String
    Dim base As String, p1 As String, p2 As String
    base = ThisWorkbook.path
    If Right$(base, 1) = "\" Or Right$(base, 1) = "/" Then
        p1 = base & fileName
        p2 = base & "templates\" & fileName
    Else
        p1 = base & "\" & fileName
        p2 = base & "\templates\" & fileName
    End If
    If Len(Dir$(p1)) > 0 Then
        ProjectTemplatePath = p1
    Else
        ProjectTemplatePath = p2
    End If
End Function


' ====================== DOWNLOADER (RAW + curl) ========================

Private Function HttpGetToFile(ByVal url As String, ByVal outPath As String) As Boolean
    On Error GoTo COMFail

    Dim http As Object
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")

    Dim redirects As Long: redirects = 0
RetryRequest:
    Debug.Print TS(), "[HTTP] GET ", url
    http.Open "GET", url, False
    http.setTimeouts 30000, 30000, 30000, 30000
    http.setRequestHeader "User-Agent", "Omexom-Orchestrator/1.0"
    http.send

    Debug.Print TS(), "[HTTP] status ", http.Status

    If http.Status = 301 Or http.Status = 302 Or http.Status = 307 Or http.Status = 308 Then
        If redirects >= 3 Then Debug.Print TS(), "[HTTP][ERROR] too many redirects": HttpGetToFile = False: Exit Function
        Dim loc As String: loc = http.getResponseHeader("Location")
        If Len(loc) = 0 Then Debug.Print TS(), "[HTTP][ERROR] redirect without Location": HttpGetToFile = False: Exit Function
        redirects = redirects + 1: url = loc
        GoTo RetryRequest
    End If

    If http.Status < 200 Or http.Status >= 300 Then
        Debug.Print TS(), "[HTTP][ERROR] non-2xx status"
        HttpGetToFile = False: Exit Function
    End If

    Dim stm As Object: Set stm = CreateObject("ADODB.Stream")
    stm.Type = 1: stm.Open
    stm.Write http.responseBody
    EnsureParentFolder outPath
    If Dir$(outPath) <> "" Then Kill outPath
    stm.SaveToFile outPath
    stm.Close

    HttpGetToFile = True
    Exit Function

COMFail:
    Debug.Print TS(), "[HTTP][WARN] COM failed (", Err.Number, ": ", Err.Description, ") -> trying curl"
    On Error GoTo CurlFail
    If CurlDownload(url, outPath) Then HttpGetToFile = True: Exit Function
    HttpGetToFile = False: Exit Function

CurlFail:
    Debug.Print TS(), "[HTTP][ERROR] curl fallback failed (", Err.Number, ": ", Err.Description, ")"
    HttpGetToFile = False
End Function

Private Function CurlDownload(ByVal url As String, ByVal outPath As String) As Boolean
    On Error GoTo EH
    EnsureParentFolder outPath

    Dim sh As Object: Set sh = CreateObject("WScript.Shell")
    Dim line As String, cmd As String
    line = "curl -L -f --silent --show-error """ & url & """ -o """ & outPath & """"

    If IsWindowsOS Then
        cmd = "cmd /c " & line
    Else
        cmd = "bash -lc """ & line & """"
    End If

    Debug.Print TS(), "[HTTP] RUN ", cmd
    CurlDownload = (sh.Run(cmd, 0, True) = 0)
    Exit Function
EH:
    Debug.Print TS(), "[HTTP][ERROR] curl not available? ", Err.Number, " - ", Err.Description
    CurlDownload = False
End Function


' ========================= FILESYSTEM HELPERS ==========================

Private Function IsWindowsOS() As Boolean
    On Error Resume Next
    IsWindowsOS = (InStr(1, Application.OperatingSystem, "Windows", vbTextCompare) > 0)
End Function

Private Function SepChar() As String
    SepChar = IIf(IsWindowsOS, "\", "/")
End Function

Private Function GetDownloadBase() As String
    Dim home As String
    If IsWindowsOS Then
        home = Environ$("USERPROFILE"): If Len(home) = 0 Then home = Environ$("HOMEPATH")
        If Len(home) = 0 Then home = "C:\Users\Public"
        GetDownloadBase = home & "\Downloads\omexom"
    Else
        home = Environ$("HOME"): If Len(home) = 0 Then home = "/tmp"
        GetDownloadBase = home & "/Downloads/omexom"
    End If
End Function

Private Function BuildOutputPath(ByVal relativePath As String) As String
    Dim base As String, sep As String
    base = GetDownloadBase(): sep = SepChar()
    If Right$(base, 1) <> sep Then base = base & sep
    BuildOutputPath = base & relativePath
End Function

Private Sub EnsureParentFolder(ByVal filePath As String)
    Dim p As String: p = ParentPath(filePath)
    If Len(p) > 0 Then MkDirs p
End Sub

Private Function ParentPath(ByVal filePath As String) As String
    Dim sep As String: sep = SepChar()
    Dim pos As Long: pos = InStrRev(filePath, sep)
    ParentPath = IIf(pos > 0, Left$(filePath, pos - 1), "")
End Function

Private Sub MkDirs(ByVal path As String)
    Dim sep As String: sep = SepChar()
    Dim parts() As String, cur As String
    Dim i As Long
    parts = Split(path, sep): If UBound(parts) < 0 Then Exit Sub

    If IsWindowsOS Then
        cur = parts(0)
        For i = 1 To UBound(parts)
            If Len(cur) > 0 And Right$(cur, 1) <> sep Then cur = cur & sep
            cur = cur & parts(i)
            If Len(Dir$(cur, vbDirectory)) = 0 Then On Error Resume Next: MkDir cur: On Error GoTo 0
        Next
    Else
        cur = IIf(Left$(path, 1) = "/", "/", "")
        For i = IIf(cur = "/", 1, 0) To UBound(parts)
            If parts(i) <> "" Then
                If Len(cur) > 0 And Right$(cur, 1) <> sep Then cur = cur & sep
                cur = cur & parts(i)
                If Len(Dir$(cur, vbDirectory)) = 0 Then On Error Resume Next: MkDir cur: On Error GoTo 0
            End If
        Next
    End If
End Sub

Private Function WorkbookPathJoin(ByVal rel As String) As String
    Dim base As String
    base = ThisWorkbook.path
    If Right$(base, 1) = "\" Or Right$(base, 1) = "/" Then
        WorkbookPathJoin = base & rel
    Else
        WorkbookPathJoin = base & "\" & rel
    End If
End Function

Private Sub WriteAllText(ByVal filePath As String, ByVal text As String)
    Dim stm As Object: Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2 ' text
    stm.Charset = "utf-8"
    stm.Open
    stm.WriteText text
    EnsureParentFolder filePath
    If Dir$(filePath) <> "" Then Kill filePath
    stm.SaveToFile filePath
    stm.Close
End Sub

Private Function ReadAllText(ByVal filePath As String) As String
    Dim stm As Object: Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2: stm.Charset = "utf-8": stm.Open
    stm.LoadFromFile filePath
    ReadAllText = stm.ReadText
    stm.Close
End Function

Private Function LoadConfigJson() As String
    Dim p As String: p = WorkbookPathJoin(CFG_REL_PATH)
    If Len(Dir$(p)) = 0 Then
        Debug.Print TS(), "[CFG][WARN] Missing config: ", p
        LoadConfigJson = ""
    Else
        LoadConfigJson = ReadAllText(p)
        Debug.Print TS(), "[CFG] Loaded: ", p
    End If
End Function

Private Function TS() As String
    TS = Format$(Now, "yyyy-mm-dd hh:nn:ss")
End Function


' =========================== TINY JSON PARSER ==========================
' NOTE: Minimal non-robust parser for flat values and simple array of objects
' Supported:
'   - JsonGetString(json, "a.b", default)
'   - JsonGetBool(json, "a.b", default)
'   - JsonGetDownloads(json, urls(), saves())  where downloads is:
'       "downloads":[{"url":"...","save_as":"..."}, ...]
' Assumptions:
'   - Double quotes around keys/values
'   - No escaped quotes inside values
'   - Whitespace not significant

Private Function JsonGetString(ByVal js As String, ByVal dottedKey As String, ByVal defaultValue As String) As String
    ' Resolve "a.b.c" to find "c" under that path by naive scanning
    Dim parts() As String: parts = Split(dottedKey, ".")
    Dim cur As String: cur = js
    Dim i As Long

    For i = 0 To UBound(parts)
        cur = JsonFindKeyBlock(cur, parts(i))
        If Len(cur) = 0 Then
            JsonGetString = defaultValue
            Exit Function
        End If
    Next i

    ' At last block, try to read a string value
    Dim v As String: v = JsonReadStringValue(cur)
    If v = "#__NOT_FOUND__#" Then
        JsonGetString = defaultValue
    Else
        JsonGetString = v
    End If
End Function

Private Function JsonGetBool(ByVal js As String, ByVal dottedKey As String, ByVal defaultValue As Boolean) As Boolean
    Dim parts() As String: parts = Split(dottedKey, ".")
    Dim cur As String: cur = js
    Dim i As Long

    For i = 0 To UBound(parts)
        cur = JsonFindKeyBlock(cur, parts(i))
        If Len(cur) = 0 Then
            JsonGetBool = defaultValue
            Exit Function
        End If
    Next i

    Dim v As String: v = JsonReadRawValue(cur)
    v = LCase$(Trim$(v))
    If v Like "true*" Then
        JsonGetBool = True
    ElseIf v Like "false*" Then
        JsonGetBool = False
    Else
        JsonGetBool = defaultValue
    End If
End Function

Private Function JsonGetDownloads(ByVal js As String, ByRef urls() As String, ByRef saves() As String) As Long
    Dim blk As String: blk = JsonFindKeyArray(js, "downloads")
    If Len(blk) = 0 Then JsonGetDownloads = 0: Exit Function

    ' Scan for {"url":"...","save_as":"..."} objects
    Dim pos As Long, count As Long
    pos = 1: count = 0
    Do
        Dim u As String, s As String
        u = JsonFindPropInObject(blk, "url", pos)
        If u = "#__END__#" Then Exit Do
        s = JsonFindPropInObject(blk, "save_as", pos)
        If s = "#__END__#" Then Exit Do
        ReDim Preserve urls(count)
        ReDim Preserve saves(count)
        urls(count) = u
        saves(count) = s
        count = count + 1
    Loop

    JsonGetDownloads = count
End Function

' ---- JSON scanning helpers (naive) ----

Private Function JsonFindKeyBlock(ByVal js As String, ByVal key As String) As String
    ' Returns substring that begins at the key and includes the value that follows
    Dim pat As String
    pat = """" & key & """"
    Dim p As Long: p = InStr(1, js, pat, vbTextCompare)
    If p = 0 Then Exit Function

    ' move to ":" after key
    Dim colon As Long: colon = InStr(p + Len(pat), js, ":", vbTextCompare)
    If colon = 0 Then Exit Function

    JsonFindKeyBlock = Mid$(js, colon + 1)
End Function

Private Function JsonFindKeyArray(ByVal js As String, ByVal key As String) As String
    Dim blk As String: blk = JsonFindKeyBlock(js, key)
    If Len(blk) = 0 Then Exit Function

    ' find [ ... ] from blk
    Dim lb As Long: lb = InStr(1, blk, "[")
    If lb = 0 Then Exit Function
    Dim depth As Long: depth = 1
    Dim i As Long
    For i = lb + 1 To Len(blk)
        Dim ch As String: ch = Mid$(blk, i, 1)
        If ch = "[" Then depth = depth + 1
        If ch = "]" Then
            depth = depth - 1
            If depth = 0 Then
                JsonFindKeyArray = Mid$(blk, lb, i - lb + 1)
                Exit Function
            End If
        End If
    Next i
End Function

Private Function JsonReadStringValue(ByVal blk As String) As String
    ' expects value like "something"
    Dim q1 As Long: q1 = InStr(1, blk, """")
    If q1 = 0 Then JsonReadStringValue = "#__NOT_FOUND__#": Exit Function
    Dim q2 As Long: q2 = InStr(q1 + 1, blk, """")
    If q2 = 0 Then JsonReadStringValue = "#__NOT_FOUND__#": Exit Function
    JsonReadStringValue = Mid$(blk, q1 + 1, q2 - q1 - 1)
End Function

Private Function JsonReadRawValue(ByVal blk As String) As String
    ' returns raw token up to comma/brace
    Dim i As Long
    Dim s As String: s = LTrim$(blk)
    For i = 1 To Len(s)
        Dim ch As String: ch = Mid$(s, i, 1)
        If ch = "," Or ch = "}" Or ch = "]" Or ch = vbCr Or ch = vbLf Then
            JsonReadRawValue = Trim$(Left$(s, i - 1))
            Exit Function
        End If
    Next i
    JsonReadRawValue = Trim$(s)
End Function

Private Function JsonFindPropInObject(ByVal arrBlk As String, ByVal prop As String, ByRef startPos As Long) As String
    ' Search from startPos for next { ... } and read the specified "prop":"value"
    Dim pObj As Long: pObj = InStr(startPos, arrBlk, "{")
    If pObj = 0 Then JsonFindPropInObject = "#__END__#": Exit Function
    Dim depth As Long: depth = 1
    Dim i As Long
    For i = pObj + 1 To Len(arrBlk)
        Dim ch As String: ch = Mid$(arrBlk, i, 1)
        If ch = "{" Then depth = depth + 1
        If ch = "}" Then
            depth = depth - 1
            If depth = 0 Then
                ' we have object substring
                Dim obj As String: obj = Mid$(arrBlk, pObj, i - pObj + 1)
                Dim blk As String: blk = JsonFindKeyBlock(obj, prop)
                If Len(blk) = 0 Then
                    JsonFindPropInObject = ""
                Else
                    JsonFindPropInObject = JsonReadStringValue(blk)
                End If
                startPos = i + 1
                Exit Function
            End If
        End If
    Next i
    JsonFindPropInObject = "#__END__#"
End Function


