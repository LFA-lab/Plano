Attribute VB_Name = "Module1"
Option Explicit

'=============================================================
' Main process : updates cell B5 on sheet "Feuil1" step by step
'=============================================================
Public Sub RunGenerationProcess()
    Dim ws As Worksheet
    Dim ole As OLEObject
    Dim shp As Shape
    Dim buttonsDisabled As Boolean

    On Error GoTo ErrorHandler

    ' <<< change the sheet name here if your tab is not "Feuil1" >>>
    Set ws = ThisWorkbook.Sheets("Feuil1")

    Debug.Print "=== PROCESS START ==="
    buttonsDisabled = False

    '---- show initial status
    ws.Range("B5").Value = "? Processing..."

    '---- disable any ActiveX CommandButtons on that sheet
    For Each ole In ws.OLEObjects
        On Error Resume Next
        If Not ole.Object Is Nothing Then
            If TypeName(ole.Object) = "CommandButton" Then
                ole.Object.Enabled = False
                buttonsDisabled = True
            End If
        End If
        On Error GoTo ErrorHandler
    Next ole

    '---- disable any Form-control buttons on that sheet
    For Each shp In ws.Shapes
        On Error Resume Next
        If shp.Type = msoFormControl Then
            If shp.FormControlType = xlButtonControl Then
                shp.ControlFormat.Enabled = False
                buttonsDisabled = True
            End If
        End If
        On Error GoTo ErrorHandler
    Next shp

    '---- demo steps – replace with your real work as needed
    ws.Range("B5").Value = "Step 1: Initializing..."
    DoEvents
    ws.Range("B5").Value = "Step 2: Processing..."
    DoEvents
    ws.Range("B5").Value = "Step 3: Finalizing..."
    DoEvents

    '---- finished
    ws.Range("B5").Value = "Complete!"   ' <- simple text so every font shows it
    GoTo Cleanup

ErrorHandler:
    ws.Range("B5").Value = "? Error: " & Err.Description
    MsgBox "An error occurred: " & Err.Description, vbExclamation, "Error"

Cleanup:
    '---- re-enable buttons even after error
    Dim ole2 As OLEObject, shp2 As Shape
    For Each ole2 In ws.OLEObjects
        On Error Resume Next
        If Not ole2.Object Is Nothing Then
            If TypeName(ole2.Object) = "CommandButton" Then ole2.Object.Enabled = True
        End If
    Next ole2

    For Each shp2 In ws.Shapes
        On Error Resume Next
        If shp2.Type = msoFormControl Then
            If shp2.FormControlType = xlButtonControl Then shp2.ControlFormat.Enabled = True
        End If
    Next shp2

    Debug.Print "=== PROCESS END ==="
End Sub

'-------------------------------------------------------------
' Optional helper: opens the UserForm directly from Excel
'-------------------------------------------------------------
Public Sub OpenGenerator()
    UserForm1.Show
End Sub

