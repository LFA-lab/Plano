Attribute VB_Name = "PlanoCore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function DownloadsFolder() As String
    DownloadsFolder = Environ$("USERPROFILE") & "\Downloads\"
End Function

Public Sub SilentDownload(ByVal url As String, ByVal destFullPath As String)
    On Error Resume Next
    Dim xhr As Object, stm As Object

    Set xhr = CreateObject("MSXML2.XMLHTTP")
    xhr.Open "GET", url, False
    xhr.send

    If xhr.Status = 200 Then
        Set stm = CreateObject("ADODB.Stream")
        stm.Type = 1
        stm.Open
        stm.Write xhr.responseBody
        stm.SaveToFile destFullPath, 2
        stm.Close
    End If
End Sub

Public Sub CreateProjectFromTemplate(ByVal templatePath As String, ByVal outputPath As String)
    If Len(Dir$(templatePath)) = 0 Then
        Err.Raise vbObjectError + 513, "PlanoCore", "Template missing: " & templatePath
    End If

    Application.FileNew Template:=templatePath
    If Len(outputPath) > 0 Then
        Application.FileSaveAs Name:=outputPath
    End If
End Sub

Public Sub RunImport()
    Dim filePath As String
    filePath = SelectPlanningFile()
    If Len(filePath) = 0 Then Exit Sub

    ImportDataSilent filePath
End Sub

Public Function SelectPlanningFile() As String
    Dim xl As Object
    Dim p As Variant

    Set xl = CreateObject("Excel.Application")
    xl.Visible = False

    p = xl.GetOpenFilename( _
            FileFilter:="Excel/CSV/MPP (*.xlsx;*.xlsm;*.csv;*.mpp),*.xlsx;*.xlsm;*.csv;*.mpp," & _
                        "Excel Files (*.xlsx;*.xlsm),*.xlsx;*.xlsm," & _
                        "MS Project Files (*.mpp),*.mpp," & _
                        "CSV Files (*.csv),*.csv", _
            Title:="Select planning file" _
        )

    If VarType(p) <> vbBoolean Then
        SelectPlanningFile = CStr(p)
    End If

    On Error Resume Next
    xl.Quit
    Set xl = Nothing
End Function

Public Sub ImportDataSilent(ByVal filePath As String)
    On Error GoTo ImportFailed

    Dim ext As String, iDot As Long
    iDot = InStrRev(filePath, ".")
    If iDot > 0 Then ext = LCase$(Mid$(filePath, iDot + 1))

    Select Case ext
        Case "mpp"
            Application.FileOpenEx Name:=filePath, ReadOnly:=False

        Case "xlsx", "xlsm", "csv"
            Dim templatePath As String
            templatePath = Application.TemplatesPath & "ModeleImport.mpt"
            CreateProjectFromTemplate templatePath, ""

        Case Else
            ' Unknown -> silent
    End Select
    Exit Sub

ImportFailed:
    MsgBox "Import failed: " & Err.Description, vbExclamation, "Plano"
End Sub
