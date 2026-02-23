Attribute VB_Name = "modPlanoMenu"
Option Explicit

Public Const PLANO_MENU_CAPTION As String = "Plano"
Public Const PLANO_MENU_TAG As String = "plano_menu_tag"

Public Const MACRO_DASHBOARD As String = "GenerateDashboard"
Public Const MACRO_IMPORT As String = "Import_Taches_Simples_AvecTitre"
Public Const MACRO_EXPORT As String = "ExportData"
Public Const MACRO_PANEL As String = "ShowPlanoControl"
Public Sub CreatePlanoMenu()
    On Error GoTo ErrHandler
    
    RemovePlanoMenu

    Dim cb As CommandBar
    Dim pop As CommandBarPopup
    
    ' --- TRY THE KNOWN COMMANDBARS IN ORDER ---
    On Error Resume Next
    ' 1) Legacy Menu Bar (common in Project Pro 2016/2019/2021)
    Set cb = Application.CommandBars("Menu Bar")
    
    ' 2) If your build doesn’t expose a visible Menu Bar ? try Menu Commands
    If cb Is Nothing Then
        Set cb = Application.CommandBars("Menu Commands")
    End If
    
    ' 3) If still nothing ? use Ribbon Adapter (Project versions with hidden Add-ins)
    If cb Is Nothing Then
        Set cb = Application.CommandBars("Ribbon")
    End If
    On Error GoTo ErrHandler

    ' --- IF NOTHING FOUND, ABORT QUIETLY ---
    If cb Is Nothing Then Exit Sub

    ' --- BUILD PLANO MENU ---
    Set pop = cb.Controls.Add(Type:=msoControlPopup, Temporary:=True)
    pop.caption = PLANO_MENU_CAPTION
    pop.Tag = PLANO_MENU_TAG

    AddPlanoButton pop, "Generate Dashboard", MACRO_DASHBOARD, 5716
    AddPlanoButton pop, "Import from Excel", MACRO_IMPORT, 19
    AddPlanoButton pop, "Export", MACRO_EXPORT, 3

    Dim sep As CommandBarButton
    Set sep = pop.Controls.Add(Type:=msoControlButton, Temporary:=True)
    sep.BeginGroup = True
    sep.Visible = False

    AddPlanoButton pop, "Open Control Panel…", MACRO_PANEL, 1086

    Exit Sub

ErrHandler:
    ' Fail silently (no UI errors)
End Sub
' ---------- Remove Menu ----------
Public Sub RemovePlanoMenu()
    On Error Resume Next
    
    Dim cb As CommandBar
    Dim ctl As CommandBarControl

    ' Remove from Menu Bar
    Set cb = Application.CommandBars("Menu Bar")
    If Not cb Is Nothing Then
        For Each ctl In cb.Controls
            If ctl.Tag = PLANO_MENU_TAG Then ctl.Delete
        Next ctl
    End If

    ' Remove from Menu Commands
    Set cb = Application.CommandBars("Menu Commands")
    If Not cb Is Nothing Then
        For Each ctl In cb.Controls
            If ctl.Tag = PLANO_MENU_TAG Then ctl.Delete
        Next ctl
    End If

    ' Remove from Ribbon fallback
    Set cb = Application.CommandBars("Ribbon")
    If Not cb Is Nothing Then
        For Each ctl In cb.Controls
            If ctl.Tag = PLANO_MENU_TAG Then ctl.Delete
        Next ctl
    End If
End Sub

' ---------- Helper ----------
Private Sub AddPlanoButton(parent As CommandBarPopup, caption As String, macroName As String, faceId As Long)
    Dim btn As CommandBarButton
    Set btn = parent.Controls.Add(Type:=msoControlButton, Temporary:=True)
    btn.caption = caption
    btn.OnAction = macroName
    btn.faceId = faceId
    btn.Style = msoButtonIconAndCaption
    btn.Tag = PLANO_MENU_TAG
End Sub

Public Sub ExportData()
  ExportToJsonModule.ExportToJson

End Sub

Sub GenerateDashboard()
    MsgBox "Button Works"
    ExportDashboardMecaElecModule.ExportDashboardMecaElec
    
End Sub





