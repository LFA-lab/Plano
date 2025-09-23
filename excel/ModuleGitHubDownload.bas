Attribute VB_Name = "Module2"
Option Explicit

' ================= Cross-platform GitHub RAW downloader =================
' - Works on Windows and Ubuntu/Wine
' - ServerXMLHTTP with redirects, timeouts; curl fallback
' - Cross-platform paths: ~/Downloads/omexom or %USERPROFILE%\Downloads\omexom
' - Includes: README test, full-URL test, and .MPP download button handler
' =======================================================================

' ---------- Configure your URLs here ----------
' For the README sanity test (always present on main):
Private Const BASE_URL As String = "https://raw.githubusercontent.com/lfa-lab/Omexom/main/"

' For .MPP downloads kept under templates/ on main:
Private Const BASE_URL_TEMPLATES As String = "https://raw.githubusercontent.com/lfa-lab/Omexom/main/templates/"

' If you prefer to point directly to one file, put the full Raw URL here
' and leave RELATIVE_MPP = "" (the button handler will use FULL_URL_MPP).
Private Const FULL_URL_MPP As String = ""   ' e.g. "https://raw.githubusercontent.com/lfa-lab/Omexom/main/templates/TemplateProject_v1.mpp"

' If you want to use BASE_URL_TEMPLATES instead, set the relative name here:
Private Const RELATIVE_MPP As String = "TemplateProject_v1.mpp"


' ----------------------------- OS helpers ------------------------------
Private Function IsWindows() As Boolean
    On Error Resume Next
    IsWindows = (InStr(1, Application.OperatingSystem, "Windows", vbTextCompare) > 0)
End Function

Private Function SepChar() As String
    SepChar = IIf(IsWindows(), "\", "/")
End Function


' -------------------------- Path utilities -----------------------------
Private Function GetDownloadBase() As String
    Dim home As String
    If IsWindows() Then
        home = Environ$("USERPROFILE")
        If Len(home) = 0 Then home = Environ$("HOMEPATH")
        If Len(home) = 0 Then home = "C:\Users\Public"
        GetDownloadBase = home & "\Downloads\omexom"
    Else
        home = Environ$("HOME")
        If Len(home) = 0 Then home = "/tmp"
        GetDownloadBase = home & "/Downloads/omexom"
    End If
End Function

Private Function BuildOutputPath(ByVal relativePath As String) As String
    Dim base As String, sep As String
    base = GetDownloadBase()
    sep = SepChar()
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
    Dim parts() As String, i As Long, cur As String

    parts = Split(path, sep)
    If UBound(parts) < 0 Then Exit Sub

    If IsWindows() Then
        cur = parts(0)                    ' e.g. "C:"
        For i = 1 To UBound(parts)
            If Len(cur) > 0 And Right$(cur, 1) <> sep Then cur = cur & sep
            cur = cur & parts(i)
            If Len(Dir$(cur, vbDirectory)) = 0 Then
                On Error Resume Next: MkDir cur: On Error GoTo 0
            End If
        Next
    Else
        cur = IIf(Left$(path, 1) = "/", "/", "")
        For i = IIf(cur = "/", 1, 0) To UBound(parts)
            If parts(i) <> "" Then
                If Len(cur) > 0 And Right$(cur, 1) <> sep Then cur = cur & sep
                cur = cur & parts(i)
                If Len(Dir$(cur, vbDirectory)) = 0 Then
                    On Error Resume Next: MkDir cur: On Error GoTo 0
                End If
            End If
        Next
    End If
End Sub


' -------------------------- HTTP + fallback ----------------------------
Private Function HttpGetToFile(ByVal url As String, ByVal outPath As String) As Boolean
    On Error GoTo COMFail

    Dim http As Object                            ' MSXML2.ServerXMLHTTP.6.0
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")

    Dim redirects As Long: redirects = 0
RetryRequest:
    Log "HTTP GET -> " & url
    http.Open "GET", url, False
    http.setTimeouts 30000, 30000, 30000, 30000   ' resolve, connect, send, receive
    http.setRequestHeader "User-Agent", "Omexom-Downloader/1.0"
    http.send

    Log "HTTP status: " & http.Status

    If http.Status = 301 Or http.Status = 302 Or http.Status = 307 Or http.Status = 308 Then
        If redirects >= 3 Then
            Log "ERROR: too many redirects": HttpGetToFile = False: Exit Function
        End If
        Dim loc As String: loc = http.getResponseHeader("Location")
        If Len(loc) = 0 Then
            Log "ERROR: redirect without Location": HttpGetToFile = False: Exit Function
        End If
        redirects = redirects + 1
        url = loc
        GoTo RetryRequest
    End If

    If http.Status < 200 Or http.Status >= 300 Then
        Log "ERROR: non-2xx status": HttpGetToFile = False: Exit Function
    End If

    Dim stm As Object: Set stm = CreateObject("ADODB.Stream") ' ADODB.Stream
    stm.Type = 1: stm.Open
    stm.Write http.responseBody
    If Dir$(outPath) <> "" Then Kill outPath
    stm.SaveToFile outPath
    stm.Close

    HttpGetToFile = True
    Exit Function

COMFail:
    Log "WARN: COM HTTP failed (" & Err.Number & ": " & Err.Description & "), trying curl"
    On Error GoTo CurlFail
    If CurlDownload(url, outPath) Then
        HttpGetToFile = True: Exit Function
    End If
    HttpGetToFile = False: Exit Function

CurlFail:
    Log "ERROR: curl fallback failed (" & Err.Number & ": " & Err.Description & ")"
    HttpGetToFile = False
End Function

Private Function CurlDownload(ByVal url As String, ByVal outPath As String) As Boolean
    On Error Go To EH
    EnsureParentFolder outPath

    Dim sh As Object: Set sh = CreateObject("WScript.Shell")
    Dim line As String, cmd As String
    line = "curl -L -f --silent --show-error """ & url & """ -o """ & outPath & """"

    If IsWindows() Then
        cmd = "cmd /c " & line
    Else
        cmd = "bash -lc """ & line & """"
    End If

    Log "RUN: " & cmd
    Dim rc As Long: rc = sh.Run(cmd, 0, True)
    CurlDownload = (rc = 0)
    Exit Function
EH:
    Log "ERROR: curl not available? " & Err.Number & " - " & Err.Description
    CurlDownload = False
End Function


' ------------------------------ Logging --------------------------------
Private Sub Log(ByVal msg As String)
    Debug.Print Format$(Now, "yyyy-mm-dd hh:nn:ss"); " | "; msg
End Sub


' ------------------------------ Tests ----------------------------------
' Shows the path the module would use to save files.
Public Sub Test_ShowPaths()
    Dim p As String
    p = BuildOutputPath("hello.txt")
    Log "OS=" & IIf(IsWindows(), "Windows", "Linux") & "  Base=" & GetDownloadBase()
    Log "Example file would save to: " & p
End Sub

' Guaranteed test: downloads README.md from main.
Public Sub Test_Download_Readme()
    Dim url As String, outPath As String
    url = BASE_URL & "README.md"
    outPath = BuildOutputPath("README.md")

    Log "=== TEST START (README) ==="
    Log "URL: " & url
    Log "OUT: " & outPath

    EnsureParentFolder outPath
    If HttpGetToFile(url, outPath) Then
        Log "SUCCESS: downloaded -> " & outPath
    Else
        Log "FAILED: could not download"
    End If
    Log "=== TEST END (README) ==="
End Sub

' One-shot: paste any full Raw/Pages URL to test directly.
Public Sub Test_FullUrl_Once()
    Dim fullUrl As String, outPath As String
    fullUrl = "https://raw.githubusercontent.com/lfa-lab/Omexom/feature/github-downloader/github-pages/Importsimple.bas"  ' <- replace when needed
    outPath = BuildOutputPath("Importsimple.bas")
    EnsureParentFolder outPath
    If HttpGetToFile(fullUrl, outPath) Then
        Log "OK -> " & outPath
    Else
        Log "FAIL"
    End If
End Sub


' ======================= .MPP DOWNLOAD BUTTON ===========================
' Button handler: downloads a .MPP and reports status in Feuil1!B5.

Public Sub DownloadProjectMpp()
    On Error GoTo EH

    Dim ws As Worksheet
    Dim url As String, outPath As String, fileName As String

    ' Choose URL & filename
    If Len(FULL_URL_MPP) > 0 Then
        fileName = GetFileNameFromUrl(FULL_URL_MPP)
        url = FULL_URL_MPP
    Else
        fileName = RELATIVE_MPP
        url = BASE_URL_TEMPLATES & RELATIVE_MPP
    End If

    ' Status cell – change sheet name/cell if needed
    Set ws = ThisWorkbook.Sheets("Feuil1")       ' <--- change if your sheet name differs
    ws.Range("B5").Value = "? downloading " & fileName & " ..."

    outPath = BuildOutputPath(fileName)
    EnsureParentFolder outPath

    If HttpGetToFile(url, outPath) Then
        ws.Range("B5").Value = "? downloaded: " & outPath
        Log "SUCCESS: .mpp downloaded -> " & outPath
    Else
        ws.Range("B5").Value = "? download failed (see Immediate window)"
        Log "FAILED: .mpp download " & url
    End If
    Exit Sub

EH:
    Log "ERROR DownloadProjectMpp: " & Err.Number & " - " & Err.Description
    On Error Resume Next
    ThisWorkbook.Sheets("Feuil1").Range("B5").Value = "? error: " & Err.Description
End Sub

Private Function GetFileNameFromUrl(ByVal u As String) As String
    Dim p As Long
    p = InStrRev(u, "/")
    If p > 0 Then
        GetFileNameFromUrl = Mid$(u, p + 1)
    Else
        GetFileNameFromUrl = u
    End If
End Function
' ===================== end of ModuleGitHubDownload ======================

