Attribute VB_Name = "Module2"
Option Explicit

' ================= GITHUB RAW FILE DOWNLOADER (robust) =================
' Downloads files directly from GitHub’s raw file host
' - Uses ServerXMLHTTP with redirect handling and timeouts
' - Falls back to curl if COM HTTP fails
' - Saves to: C:\Users\<you>\Downloads\omexom\<file>
' ========================================================================

' ? BASE URL: raw file host for branch "main"
'    (change "main" if your files are on a different branch)
Private Const BASE_URL As String = "https://raw.githubusercontent.com/lfa-lab/Omexom/main/github-pages/"

' -------- PUBLIC: quick test --------
Public Sub Test_Download_XmlHttp()
    Dim relative As String
    relative = "Importsimple.bas"  ' exact file name under github-pages/

    Dim url As String, outPath As String
    url = BASE_URL & relative
    outPath = BuildOutputPath(relative)

    Log "=== TEST START ==="
    Log "URL: " & url
    Log "OUT: " & outPath

    EnsureParentFolder outPath
    If HttpGetToFile(url, outPath) Then
        Log "SUCCESS: downloaded -> " & outPath
    Else
        Log "FAILED: could not download"
    End If
    Log "=== TEST END ==="
End Sub

' -------- HTTP with redirects + curl fallback --------
Private Function HttpGetToFile(ByVal url As String, ByVal outPath As String) As Boolean
    On Error GoTo COMFail

    Dim http As Object ' MSXML2.ServerXMLHTTP.6.0
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")

    Dim redirects As Long: redirects = 0
RetryRequest:
    Log "HTTP GET -> " & url
    http.Open "GET", url, False
    http.setTimeouts 30000, 30000, 30000, 30000   ' resolve, connect, send, receive (ms)
    http.setRequestHeader "User-Agent", "Omexom-Downloader/1.0"
    http.send

    Log "HTTP status: " & http.Status

    ' follow redirects (GitHub raw may redirect)
    If http.Status = 301 Or http.Status = 302 Or http.Status = 307 Or http.Status = 308 Then
        If redirects >= 3 Then
            Log "ERROR: too many redirects"
            HttpGetToFile = False: Exit Function
        End If
        Dim loc As String: loc = http.getResponseHeader("Location")
        If Len(loc) = 0 Then
            Log "ERROR: redirect without Location"
            HttpGetToFile = False: Exit Function
        End If
        redirects = redirects + 1
        url = loc
        GoTo RetryRequest
    End If

    If http.Status < 200 Or http.Status >= 300 Then
        Log "ERROR: non-2xx status"
        HttpGetToFile = False: Exit Function
    End If

    ' save binary
    Dim stm As Object ' ADODB.Stream
    Set stm = CreateObject("ADODB.Stream")
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

' -------- curl fallback (Win10+/Linux usually has curl) --------
Private Function CurlDownload(ByVal url As String, ByVal outPath As String) As Boolean
    On Error GoTo EH
    EnsureParentFolder outPath
    Dim sh As Object: Set sh = CreateObject("WScript.Shell")
    Dim cmd As String
    cmd = "curl -L -f --silent --show-error """ & url & """ -o """ & outPath & """"
    Log "RUN: " & cmd
    Dim rc As Long: rc = sh.Run(cmd, 0, True)
    CurlDownload = (rc = 0)
    Exit Function
EH:
    Log "ERROR: curl not available? " & Err.Number & " - " & Err.Description
    CurlDownload = False
End Function

' -------- paths / folders (Windows step) --------
Private Function BuildOutputPath(ByVal relativePath As String) As String
    Dim base As String: base = GetDownloadBase()
    If Right$(base, 1) <> "\" Then base = base & "\"
    BuildOutputPath = base & relativePath
End Function

Private Function GetDownloadBase() As String
    Dim home As String
    home = Environ$("USERPROFILE"): If Len(home) = 0 Then home = Environ$("HOMEPATH")
    If Len(home) = 0 Then home = "C:\Users\Public"
    GetDownloadBase = home & "\Downloads\omexom"
End Function

Private Sub EnsureParentFolder(ByVal filePath As String)
    Dim p As String: p = ParentPath(filePath)
    If Len(p) > 0 Then MkDirs p
End Sub

Private Function ParentPath(ByVal filePath As String) As String
    Dim pos As Long: pos = InStrRev(filePath, "\")
    ParentPath = IIf(pos > 0, Left$(filePath, pos - 1), "")
End Function

Private Sub MkDirs(ByVal path As String)
    Dim parts() As String, i As Long, cur As String
    parts = Split(path, "\")
    If UBound(parts) < 0 Then Exit Sub
    cur = parts(0)
    For i = 1 To UBound(parts)
        cur = cur & "\" & parts(i)
        If Len(Dir$(cur, vbDirectory)) = 0 Then
            On Error Resume Next: MkDir cur: On Error GoTo 0
        End If
    Next
End Sub

' -------- logging --------
Private Sub Log(ByVal msg As String)
    Debug.Print Format$(Now, "yyyy-mm-dd hh:nn:ss"); " | "; msg
End Sub

