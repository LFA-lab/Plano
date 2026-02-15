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
    ' Branding per senior�s request; keep UX silent
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
    url = "https://lfa-lab.github.io/Plano/templates/FichierTypearemplir.xlsx"
    dest = DownloadsFolder() & "FichierTypearemplir.xlsx"

    SilentDownload url, dest
    If DEBUG_LOG Then Debug.Print "Template saved to:", dest
    ' UX: remain silent (no MsgBox)
End Sub

' ========= SMALL BUTTON: Cancel =========
Private Sub btnCancel_Click()
    Unload Me
End Sub

' ========= IMPORT CONTROLLER - CALLS Import_OPTIMISE =========
' This is the controller that:
' 1) Calls Import_OPTIMISE.Import_Taches_Simples_AvecTitre (does the Excel import)
' 2) Saves the result as .mpp
' 3) Opens the .mpp automatically
Private Sub ImportDataSilent(ByVal filePath As String)
    On Error GoTo ImportError

    Dim ext As String, iDot As Long
    iDot = InStrRev(filePath, ".")
    If iDot > 0 Then ext = LCase$(Mid$(filePath, iDot + 1))

    Select Case ext
        Case "mpp"
            ' Open existing .mpp silently (read/write)
            Application.FileOpenEx Name:=filePath, ReadOnly:=False

        Case "xlsx", "xlsm", "csv"
            ' ===== WORKFLOW: Import Excel → Create .mpp → Open .mpp =====

            ' STEP 1: Call Import_OPTIMISE to create the project from Excel
            ' Note: Import_Taches_Simples_AvecTitre will:
            '  - Ask user to select Excel file (already done via filePath)
            '  - Create project structure
            '  - We need to intercept this to pass our filePath

            ' STEP 2: Since Import_OPTIMISE expects to select file via dialog,
            ' we'll use a wrapper approach:
            ' - Store filePath in a temp variable
            ' - Call a modified import that uses our filePath

            ' For now, call the import and let it create the project
            ' The user has already selected the file via our dialog
            Call Import_Taches_Simples_AvecTitre_FromUserForm(filePath)

            ' STEP 3: Save as .mpp next to the Excel file
            Dim mppPath As String
            mppPath = Replace(filePath, ".xlsx", ".mpp")
            mppPath = Replace(mppPath, ".xlsm", ".mpp")
            mppPath = Replace(mppPath, ".csv", ".mpp")

            On Error Resume Next
            Application.FileSaveAs Name:=mppPath
            On Error GoTo ImportError

            If DEBUG_LOG Then Debug.Print "Project saved as:", mppPath

            ' STEP 4: The .mpp is now open and active
            ' The Project_Open event in ThisProject will detect .mpp
            ' and create the Plano menu automatically

        Case Else
            ' Unknown -> do nothing (silent)
    End Select

    If DEBUG_LOG Then Debug.Print "ImportDataSilent:", filePath, "ext=", ext
    Exit Sub

ImportError:
    ' Log error but remain silent (no MsgBox per UX mandate)
    If DEBUG_LOG Then Debug.Print "ImportDataSilent ERROR:", Err.Number, Err.Description
End Sub

' ========= WRAPPER FOR Import_OPTIMISE =========
' This sub wraps the Import_Taches_Simples_AvecTitre call
' to pass the pre-selected file path instead of showing dialog
Private Sub Import_Taches_Simples_AvecTitre_FromUserForm(ByVal preSelectedFile As String)
    ' Call the wrapper function that passes the file parameter
    On Error Resume Next
    Call Import_Taches_Simples_AvecTitre_WithFile(preSelectedFile)
End Sub
