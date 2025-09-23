Attribute VB_Name = "Module3"
Option Explicit

' ==============================================================
' MS Project COM integration (Windows) with extensive logging
' - Opens a .mpt template
' - Runs existing macros (no business macro modifications)
' - Optional "carrier project" pattern
' - Detects non-Windows and explains next actions
' ==============================================================

' -------- Public tests --------

Public Sub Test_OpenMpt()
    Dim tplPath As String
    tplPath = ProjectTemplatePath("TemplateProject_v1.mpt")
    MSP_OpenCloseTemplate tplPath, False, ""
End Sub

' Set macroName to an existing macro in the template:
' Example: "TemplateProject_v1!Module1.SampleMacro"
Public Sub Test_OpenMptAndRunMacro()
    Dim tplPath As String, macroName As String
    tplPath = ProjectTemplatePath("TemplateProject_v1.mpt")
    macroName = "TemplateProject_v1!Module1.SampleMacro" ' TODO: replace with real one
    MSP_OpenCloseTemplate tplPath, True, macroName
End Sub

' Optional carrier: a project that holds helper macros you can run on ActiveProject
Public Sub Test_RunCarrierMacro()
    Dim tplPath As String, carrierPath As String, carrierMacro As String
    tplPath = ProjectTemplatePath("TemplateProject_v1.mpt")
    carrierPath = ProjectTemplatePath("UtilitiesCarrier.mpp") ' optional carrier
    carrierMacro = "UtilitiesCarrier!Helpers.DoStuffOnActiveProject" ' TODO if used
    MSP_OpenRunCarrier tplPath, carrierPath, carrierMacro
End Sub

' -------- Core --------

Private Sub MSP_OpenCloseTemplate(ByVal templatePath As String, _
                                  ByVal runMacro As Boolean, _
                                  ByVal macroName As String)
    If Not IsWindowsOS Then
        Log "[MSP] Non-Windows OS. COM automation not available here."
        Log "[MSP] Use Windows VM or Wine launch scripts. Skipping."
        Exit Sub
    End If

    On Error GoTo EH
    Dim app As Object, proj As Object           ' late binding

    Log "[MSP] Starting MS Project..."
    Set app = CreateObject("MSProject.Application")
    app.Visible = True

    Log "[MSP] Opening template: " & templatePath
    app.FileOpen Name:=templatePath

    Set proj = app.ActiveProject
    Log "[MSP] ActiveProject: " & proj.Name

    If runMacro And Len(Trim$(macroName)) > 0 Then
        Log "[MSP] Running macro: " & macroName
        app.Run macroName
        Log "[MSP] Macro finished: " & macroName
    Else
        Log "[MSP] No macro requested (smoke test)."
    End If

    Log "[MSP] Closing without save."
    app.FileClose pjDoNotSave

    Log "[MSP] Quitting."
    app.Quit

    Set proj = Nothing
    Set app = Nothing
    Log "[MSP] DONE."
    Exit Sub

EH:
    Log "[MSP][ERROR] " & Err.Number & " - " & Err.Description
    On Error Resume Next
    If Not app Is Nothing Then
        app.FileClose pjDoNotSave
        app.Quit
    End If
End Sub

Private Sub MSP_OpenRunCarrier(ByVal templatePath As String, _
                               ByVal carrierProjectPath As String, _
                               ByVal carrierMacroName As String)
    If Not IsWindowsOS Then
        Log "[MSP] Non-Windows OS. COM automation not available here."
        Log "[MSP] Use Windows VM or Wine launch scripts. Skipping."
        Exit Sub
    End If

    On Error GoTo EH
    Dim app As Object, target As Object, carrier As Object

    Log "[MSP] Starting MS Project..."
    Set app = CreateObject("MSProject.Application")
    app.Visible = True

    Log "[MSP] Opening target template: " & templatePath
    app.FileOpen Name:=templatePath
    Set target = app.ActiveProject
    Log "[MSP] Target: " & target.Name

    Log "[MSP] Opening carrier project: " & carrierProjectPath
    app.FileOpen Name:=carrierProjectPath
    Set carrier = app.ActiveProject
    Log "[MSP] Carrier: " & carrier.Name

    app.ActivateProject target.Name

    Log "[MSP] Running carrier macro: " & carrierMacroName
    app.Run carrierMacroName
    Log "[MSP] Carrier macro finished."

    Log "[MSP] Closing projects (no save)."
    app.FileCloseAll pjDoNotSave

    Log "[MSP] Quitting."
    app.Quit

    Set carrier = Nothing
    Set target = Nothing
    Set app = Nothing
    Log "[MSP] DONE."
    Exit Sub

EH:
    Log "[MSP][ERROR] " & Err.Number & " - " & Err.Description
    On Error Resume Next
    If Not app Is Nothing Then
        app.FileCloseAll pjDoNotSave
        app.Quit
    End If
End Sub

' -------- Utilities --------

Private Function ProjectTemplatePath(ByVal fileName As String) As String
    Dim base As String
    base = ThisWorkbook.path
    If Right$(base, 1) = "\" Or Right$(base, 1) = "/" Then
        ProjectTemplatePath = base & "templates\" & fileName
    Else
        ProjectTemplatePath = base & "\templates\" & fileName
    End If
End Function

Private Function IsWindowsOS() As Boolean
    On Error Resume Next
    IsWindowsOS = (InStr(1, Application.OperatingSystem, "Windows", vbTextCompare) > 0)
End Function

Private Sub Log(ByVal msg As String)
    Debug.Print Format$(Now, "yyyy-mm-dd hh:nn:ss"); " | "; msg
End Sub


