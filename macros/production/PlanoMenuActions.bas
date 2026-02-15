Attribute VB_Name = "PlanoMenuActions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'=================================================================
' PlanoMenuActions - Menu action handlers for Plano
'=================================================================
' This module contains all the Public Sub procedures that are
' called by the Plano menu items (OnAction handlers)
'=================================================================

'=================================================================
' GenerateDashboard - Export project data and generate HTML dashboard
'=================================================================
Public Sub GenerateDashboard()
    On Error GoTo ErrorHandler

    ' Check if there's an active project
    If ActiveProject Is Nothing Then
        MsgBox "No active project. Please open a project first.", vbExclamation, "Plano"
        Exit Sub
    End If

    ' Call the ExportToJson function which generates the dashboard data
    On Error Resume Next
    ExportToJsonModule.ExportToJson
    If Err.Number <> 0 Then
        MsgBox "Error generating dashboard: " & Err.Description, vbExclamation, "Plano"
        Err.Clear
        Exit Sub
    End If
    On Error GoTo ErrorHandler

    ' Success message
    MsgBox "Dashboard data exported successfully!" & vbCrLf & vbCrLf & _
           "Open the HTML dashboard file to view your project status.", _
           vbInformation, "Plano"

    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description, vbCritical, "Plano - Generate Dashboard"
End Sub

'=================================================================
' ImportFromExcel - Import Excel data into current project
'=================================================================
Public Sub ImportFromExcel()
    On Error GoTo ErrorHandler

    ' Use Excel file picker to select file
    Dim fd As Object
    Dim f As String

    On Error Resume Next
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    If Err.Number <> 0 Then
        ' FileDialog not available, use alternative
        Err.Clear
        On Error GoTo ErrorHandler
        f = PlanoCore.SelectPlanningFile()
        If Len(f) = 0 Then Exit Sub
        PlanoCore.ImportDataSilent f
        Exit Sub
    End If
    On Error GoTo ErrorHandler

    fd.AllowMultiSelect = False
    fd.Title = "Select Excel File to Import"

    fd.Filters.Clear
    fd.Filters.Add "Excel Files", "*.xlsx;*.xlsm"

    If fd.Show <> -1 Then Exit Sub
    f = fd.SelectedItems(1)

    ' Import the file
    PlanoCore.ImportDataSilent f

    Exit Sub

ErrorHandler:
    MsgBox "An error occurred during import: " & Err.Description, vbCritical, "Plano - Import"
End Sub

'=================================================================
' ExportData - Export project data to JSON
'=================================================================
Public Sub ExportData()
    On Error GoTo ErrorHandler

    ' Check if there's an active project
    If ActiveProject Is Nothing Then
        MsgBox "No active project. Please open a project first.", vbExclamation, "Plano"
        Exit Sub
    End If

    ' Call the ExportToJson module
    On Error Resume Next
    ExportToJsonModule.ExportToJson
    If Err.Number <> 0 Then
        MsgBox "Error exporting data: " & Err.Description, vbExclamation, "Plano"
        Err.Clear
        Exit Sub
    End If
    On Error GoTo ErrorHandler

    MsgBox "Data exported successfully to JSON!", vbInformation, "Plano"

    Exit Sub

ErrorHandler:
    MsgBox "An error occurred during export: " & Err.Description, vbCritical, "Plano - Export"
End Sub

'=================================================================
' ShowPlanoControl - Show Plano control panel (future feature)
'=================================================================
Public Sub ShowPlanoControl()
    On Error Resume Next

    ' Check if UserFormPlanoControl exists
    Dim frm As Object
    Set frm = VBA.UserForms.Add("UserFormPlanoControl")

    If frm Is Nothing Then
        ' Control panel form doesn't exist yet
        MsgBox "Plano Control Panel" & vbCrLf & vbCrLf & _
               "Version: 1.0" & vbCrLf & _
               "Project: " & ActiveProject.Name & vbCrLf & vbCrLf & _
               "Features:" & vbCrLf & _
               "• Generate Dashboard - Export project status to HTML" & vbCrLf & _
               "• Import from Excel - Import tasks from Excel template" & vbCrLf & _
               "• Export to JSON - Export project data" & vbCrLf & vbCrLf & _
               "For more information, visit: https://lfa-lab.github.io/Plano/", _
               vbInformation, "Plano Control Panel"
    Else
        ' Show the control panel form
        frm.Show vbModeless
    End If
End Sub
