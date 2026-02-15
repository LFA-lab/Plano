VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisProject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'=================================================================
' ThisProject - Event Handlers for Plano Workflow
'=================================================================
' ARCHITECTURE CIBLE:
' - .mpt → Show UserFormImport automatically (NO Plano menu)
' - .mpp → Show Plano menu automatically (NO UserFormImport)
'=================================================================

Private Sub Project_Open(ByVal pj As Project)
    On Error Resume Next

    ' Detect if current file is .mpt (template) or .mpp (project)
    Dim fileName As String
    Dim fileExt As String

    fileName = ActiveProject.FullName

    ' Extract extension (last 4 characters)
    If Len(fileName) > 4 Then
        fileExt = LCase$(Right$(fileName, 4))
    Else
        fileExt = ""
    End If

    ' WORKFLOW LOGIC:
    ' If .mpt → Show UserFormImport
    ' If .mpp → Create Plano Menu

    If fileExt = ".mpt" Then
        ' ===== TEMPLATE MODE =====
        ' Show the import form automatically
        On Error Resume Next
        UserFormImport.Show vbModeless
        On Error GoTo 0

    ElseIf fileExt = ".mpp" Then
        ' ===== PROJECT MODE =====
        ' Create Plano menu (do NOT show UserFormImport)
        CreatePlanoMenu

    End If
End Sub

Private Sub Project_BeforeClose(ByVal pj As Project)
    ' Clean up Plano menu when closing project
    On Error Resume Next
    RemovePlanoMenu
    On Error GoTo 0
End Sub

'=================================================================
' CreatePlanoMenu - Build Plano menu in CommandBars
'=================================================================
' CRITICAL: OnAction must NOT contain filename (portable code)
' CORRECT:   item.OnAction = "MacroName"
' INCORRECT: item.OnAction = "File.mpt!MacroName"
'=================================================================
Public Sub CreatePlanoMenu()
    On Error GoTo ErrHandler

    ' Remove any existing Plano menu first
    RemovePlanoMenu

    Dim cb As CommandBar
    Dim pop As CommandBarPopup

    ' Try to find the menu bar (try multiple locations for compatibility)
    On Error Resume Next

    ' 1) Legacy Menu Bar (Project Pro 2016/2019/2021)
    Set cb = Application.CommandBars("Menu Bar")

    ' 2) If Menu Bar not found, try Menu Commands
    If cb Is Nothing Then
        Set cb = Application.CommandBars("Menu Commands")
    End If

    ' 3) If still nothing, try Ribbon adapter
    If cb Is Nothing Then
        Set cb = Application.CommandBars("Ribbon")
    End If

    On Error GoTo ErrHandler

    ' If no CommandBar found, exit silently
    If cb Is Nothing Then Exit Sub

    ' Create Plano menu
    Set pop = cb.Controls.Add(Type:=msoControlPopup, Temporary:=True)
    pop.Caption = "Plano"
    pop.Tag = "plano_menu_tag"

    ' Add menu items - CRITICAL: OnAction WITHOUT filename
    AddPlanoButton pop, "Generate Dashboard", "GenerateDashboard", 5716
    AddPlanoButton pop, "Import from Excel", "ImportFromExcel", 19
    AddPlanoButton pop, "Export to JSON", "ExportData", 3

    ' Separator
    Dim sep As CommandBarButton
    Set sep = pop.Controls.Add(Type:=msoControlButton, Temporary:=True)
    sep.BeginGroup = True
    sep.Visible = False

    ' Control panel
    AddPlanoButton pop, "Control Panel...", "ShowPlanoControl", 1086

    Exit Sub

ErrHandler:
    ' Fail silently - do not interrupt user workflow
End Sub

'=================================================================
' RemovePlanoMenu - Clean up Plano menu from all CommandBars
'=================================================================
Public Sub RemovePlanoMenu()
    On Error Resume Next

    Dim cb As CommandBar
    Dim ctl As CommandBarControl

    ' Remove from all possible locations
    Dim barNames As Variant
    barNames = Array("Menu Bar", "Menu Commands", "Ribbon")

    Dim barName As Variant
    For Each barName In barNames
        Set cb = Application.CommandBars(CStr(barName))
        If Not cb Is Nothing Then
            For Each ctl In cb.Controls
                If ctl.Tag = "plano_menu_tag" Then
                    ctl.Delete
                End If
            Next ctl
        End If
    Next barName

    On Error GoTo 0
End Sub

'=================================================================
' AddPlanoButton - Helper to add menu items
'=================================================================
' CRITICAL: macroName must be RELATIVE (no filename prefix)
'=================================================================
Private Sub AddPlanoButton(parent As CommandBarPopup, Caption As String, macroName As String, faceId As Long)
    Dim btn As CommandBarButton
    Set btn = parent.Controls.Add(Type:=msoControlButton, Temporary:=True)
    btn.Caption = Caption
    btn.OnAction = macroName  ' ✅ CORRECT: Relative macro name
    btn.faceId = faceId
    btn.Style = msoButtonIconAndCaption
    btn.Tag = "plano_menu_tag"
End Sub
