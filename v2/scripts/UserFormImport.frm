VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormImport 
   Caption         =   "Plano - Import Planning Data"
   ClientHeight    =   3040
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   7780
   OleObjectBlob   =   "UserFormImport.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFormImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ========= Developer toggle (no UI popups either way) =========
Private Const DEBUG_LOG As Boolean = False  ' True -> prints to Immediate Window (Ctrl+G)

' ========= Utilities: Downloads path & silent web download =========
Private Function DownloadsFolder() As String
    ' Simple default: USERPROFILE\Downloads
    DownloadsFolder = Environ$("USERPROFILE") & "\Downloads\"
End Function

Private Sub SilentDownload(ByVal url As String, ByVal destFullPath As String)
    ' 100% VBA (no Win32 Declare): MSXML2.XMLHTTP + ADODB.Stream
    ' Silent: no MsgBox, no UI
    On Error Resume Next
    Dim xhr As Object, stm As Object

    Set xhr = CreateObject("MSXML2.XMLHTTP")
    xhr.Open "GET", url, False
    xhr.send

    If xhr.Status = 200 Then
        Set stm = CreateObject("ADODB.Stream")
        stm.Type = 1 ' adTypeBinary
        stm.Open
        stm.Write xhr.responseBody
        stm.SaveToFile destFullPath, 2 ' adSaveCreateOverWrite
        stm.Close
    End If

    If DEBUG_LOG Then Debug.Print "SilentDownload:", url, "?", destFullPath, "status=", xhr.Status
End Sub

' ========= Form lifecycle =========
Private Sub UserForm_Initialize()
    ' Branding per senior’s request; keep UX silent
    Me.Caption = "Plano - Import Planning Data"
End Sub

' ========= BIG BUTTON: Browse and Select File =========
' One click ? Excel picker (.xlsx/.xlsm/.csv/.mpp) ? silent import ? form closes
Private Sub btnBrowseFile_Click()
    Dim xl As Object, p As Variant, okToProceed As Boolean
    On Error GoTo SilentCleanup

    ' Excel's file picker is robust when called from Project via late binding
    Set xl = CreateObject("Excel.Application")
    xl.Visible = False

    p = xl.GetOpenFilename( _
            FileFilter:="Excel/CSV/MPP (*.xlsx;*.xlsm;*.csv;*.mpp),*.xlsx;*.xlsm;*.csv;*.mpp," & _
                        "Excel Files (*.xlsx;*.xlsm),*.xlsx;*.xlsm," & _
                        "MS Project Files (*.mpp),*.mpp," & _
                        "CSV Files (*.csv),*.csv", _
            Title:="Select planning file" _
        )

    If VarType(p) = vbBoolean Then GoTo SilentCleanup ' user cancelled (stay silent)

    okToProceed = (Len(CStr(p)) > 0)
    If okToProceed Then
        If Dir$(CStr(p)) = "" Then okToProceed = False
    End If

    If okToProceed Then
        ImportDataSilent CStr(p)  ' no popups, no confirmations
    End If

SilentCleanup:
    On Error Resume Next
    If Not xl Is Nothing Then xl.Quit
    Set xl = Nothing
    Unload Me  ' auto-close form per UX
End Sub

' ========= SMALL BUTTON: Download Template (direct to Downloads, no browser) =========
Private Sub btnDownloadTemplate_Click()
    On Error Resume Next
    Dim url As String, dest As String
    url = "https://lfa-lab.github.io/Plano/templates/Mod%C3%A8leImport.mpt"
    dest = DownloadsFolder() & "PlanningTemplate.xlsx"

    SilentDownload url, dest
    If DEBUG_LOG Then Debug.Print "Template saved to:", dest
    ' UX: remain silent (no MsgBox)
End Sub

' ========= SMALL BUTTON: Cancel =========
Private Sub btnCancel_Click()
    Unload Me
End Sub

' ========= SILENT IMPORT CONTROLLER (stub) =========
' Implement real Excel/CSV ? MPP mapping later (still silent).
Private Sub ImportDataSilent(ByVal filePath As String)
    On Error Resume Next

    Dim ext As String, iDot As Long
    iDot = InStrRev(filePath, ".")
    If iDot > 0 Then ext = LCase$(Mid$(filePath, iDot + 1))

    Select Case ext
        Case "mpp"
            ' Open existing .mpp silently (read/write)
            Application.FileOpenEx Name:=filePath, ReadOnly:=False

        Case "xlsx", "xlsm", "csv"
            ' TODO (when mapping rules are available):
            ' 1) Open/create a Project
            ' 2) Read rows from Excel/CSV
            ' 3) Create tasks/resources/assignments
            ' 4) Save as .mpp next to source
            ' All without UI. Keep silent per UX mandate.

        Case Else
            ' Unknown -> do nothing (silent)
    End Select

    If DEBUG_LOG Then Debug.Print "ImportDataSilent:", filePath, "ext=", ext
    ' No popups. No confirmations.
End Sub
