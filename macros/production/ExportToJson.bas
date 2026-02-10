Attribute VB_Name = "ExportToJsonModule"
Option Explicit

' ============================================================
' CLEANED, STABLE, FINAL MASTER MACRO
' Export JSON aligned with senior spec + Option A aggregation
' ============================================================

Public Sub ExportToJson()
    Dim proj As Project
    Set proj = ActiveProject
    
    ' Ensure project is saved
    If Len(proj.Path) = 0 Then
        MsgBox "Please save the project (.mpp) first.", vbExclamation
        Exit Sub
    End If
    
    Dim sJson As String
    Dim sFilePath As String
    sFilePath = proj.Path & "\project_data.json"
    
    sJson = "{"
    
    ' -------------------------
    ' Header
    ' -------------------------
    sJson = sJson & """project_name"": """ & EscapeJSON(proj.Name) & ""","
    sJson = sJson & """date_export"": """ & Format(Now, "yyyy-mm-dd\Thh:nn:ss") & ""","
    
    ' -------------------------
    ' Tasks
    ' -------------------------
    sJson = sJson & """tasks"": ["
    
    Dim t As Task
    Dim firstTask As Boolean: firstTask = True
    
    For Each t In proj.Tasks
        If Not t Is Nothing Then
            If t.Summary = False Then
            
                If Not firstTask Then sJson = sJson & ","
                firstTask = False
                
                sJson = sJson & "{"
                sJson = sJson & """task_id"": " & CLng(t.ID) & ","
                sJson = sJson & """name"": """ & EscapeJSON(NzText(t.Name)) & ""","
                sJson = sJson & """percent_complete"": " & CLng(NzNumSafe(t.PercentComplete, 0)) & ","
                
                ' Physical %
                If IsNull(t.Number1) Then
                    sJson = sJson & """percent_physical"": null,"
                Else
                    sJson = sJson & """percent_physical"": " & NumToJSON(CDbl(t.Number1)) & ","
                End If
                
                sJson = sJson & """date_start_planned"": """ & DateToStr(t.Start) & ""","
                sJson = sJson & """date_finish_planned"": """ & DateToStr(t.Finish) & ""","
                
                ' Actual finish
                If IsDateValid(t.ActualFinish) Then
                    sJson = sJson & """date_finish_actual"": """ & DateToStr(CDate(t.ActualFinish)) & """"
                Else
                    sJson = sJson & """date_finish_actual"": null"
                End If
                
                sJson = sJson & "}"
            End If
        End If
    Next t
    
    sJson = sJson & "],"
    
    ' -------------------------
    ' Resources
    ' -------------------------
    sJson = sJson & """resources"": ["
    
    Dim r As Resource
    Dim firstRes As Boolean: firstRes = True
    
    For Each r In proj.Resources
        If Not r Is Nothing Then
            If Len(Trim$(r.Name)) > 0 Then
            
                ' DAILY AGGREGATION DICTIONARY
                Dim dailyDict As Object
                Set dailyDict = CreateObject("Scripting.Dictionary")
                
                Dim startD As Date: startD = proj.ProjectStart
                Dim finishD As Date: finishD = proj.ProjectFinish
                
                Dim asg As Assignment
                For Each asg In r.Assignments
                    If Not asg Is Nothing Then
                        Dim i As Long
                        
                        ' Planned
                        Dim tsP As TimeScaleValues
                        Set tsP = asg.TimeScaleData(startD, finishD, pjAssignmentTimescaledWork, pjTimescaleDays, 1)
                        
                        If Not tsP Is Nothing Then
                            For i = 1 To tsP.Count
                                AddMinutesToDaily dailyDict, tsP.Item(i).StartDate, NzNumSafe(tsP.Item(i).Value, 0), True
                            Next i
                        End If
                        
                        ' Actual
                        Dim tsA As TimeScaleValues
                        Set tsA = asg.TimeScaleData(startD, finishD, pjAssignmentTimescaledActualWork, pjTimescaleDays, 1)
                        
                        If Not tsA Is Nothing Then
                            For i = 1 To tsA.Count
                                AddMinutesToDaily dailyDict, tsA.Item(i).StartDate, NzNumSafe(tsA.Item(i).Value, 0), False
                            Next i
                        End If
                        
                    End If
                Next asg
                
                ' Emit resource block
                If Not firstRes Then sJson = sJson & ","
                firstRes = False
                
                sJson = sJson & "{"
                sJson = sJson & """resource_name"": """ & EscapeJSON(r.Name) & ""","
                sJson = sJson & """resource_type"": """ & ResourceTypeName(r.Type) & ""","
                
                ' Build sorted date list safely using ArrayList
                Dim dateKeys As Variant
                dateKeys = dailyDict.Keys  ' This is a Variant array

                ' Sort the keys safely
                If IsArray(dateKeys) Then
                    Call SortVariantArray(dateKeys)
                End If
                
                ' DAILY DATA OUTPUT
                sJson = sJson & """daily_data"": ["
                
                Dim firstDay As Boolean: firstDay = True
                Dim idx As Long
                Dim dayKey As String
                For idx = LBound(dateKeys) To UBound(dateKeys)
                    dayKey = CStr(dateKeys(idx))
               
                    
                    Dim plannedH As Double, actualH As Double
                    plannedH = dailyDict(dayKey)("p")
                    actualH = dailyDict(dayKey)("a")
                    
                    If Not firstDay Then sJson = sJson & ","
                    firstDay = False
                    
                    sJson = sJson & "{"
                    sJson = sJson & """date"": """ & dayKey & ""","
                    sJson = sJson & """qty_planned"": " & NumToJSON(plannedH) & ","
                    sJson = sJson & """qty_actual"": " & NumToJSON(actualH) & ","
                    
                    If plannedH > 0 Then
                        sJson = sJson & """percent_done"": " & NumToJSON((actualH / plannedH) * 100#)
                    Else
                        sJson = sJson & """percent_done"": null"
                    End If
                    
                    sJson = sJson & "}"
                Next idx
                
                sJson = sJson & "]" ' end daily_data
                
                sJson = sJson & "}"  ' end resource
            End If
        End If
    Next r
    
    sJson = sJson & "]"   ' end resources
    sJson = sJson & "}"   ' end root
    
    ' PRETTY PRINT JSON
    sJson = PrettyJSON(sJson, 2)
    
    ' WRITE FILE
    Dim f As Integer
    On Error GoTo FILE_WRITE_ERR
    f = FreeFile
    Open sFilePath For Output As #f
    Print #f, sJson
    Close #f
    On Error GoTo 0
    
    MsgBox "JSON generated: " & sFilePath, vbInformation
    Exit Sub

FILE_WRITE_ERR:
    MsgBox "Failed to write JSON: " & Err.Description, vbCritical
End Sub

' ============================================================
' HELPERS
' ============================================================

Private Sub AddMinutesToDaily(ByRef dict As Object, ByVal d As Date, ByVal mins As Double, ByVal isPlanned As Boolean)
    Dim key As String: key = Format(d, "yyyy-mm-dd")
    
    If Not dict.Exists(key) Then
        Dim entry As Object
        Set entry = CreateObject("Scripting.Dictionary")
        entry.Add "p", 0#
        entry.Add "a", 0#
        dict.Add key, entry
    End If
    
    If isPlanned Then
        dict(key)("p") = dict(key)("p") + (mins / 60#)
    Else
        dict(key)("a") = dict(key)("a") + (mins / 60#)
    End If
End Sub

Private Function EscapeJSON(text As String) As String
    If Len(text) = 0 Then
        EscapeJSON = ""
        Exit Function
    End If
    text = Replace(text, "\", "\\")
    text = Replace(text, """", "\""")
    EscapeJSON = text
End Function

Private Function NzText(v As Variant) As String
    If IsNull(v) Then NzText = "" Else NzText = CStr(v)
End Function

Private Function NzNumSafe(v As Variant, ByVal def As Double) As Double
    On Error GoTo FAIL
    If IsNull(v) Then
        NzNumSafe = def
    ElseIf IsNumeric(v) Then
        NzNumSafe = CDbl(v)
    Else
FAIL:
        NzNumSafe = def
    End If
End Function

Private Function IsDateValid(v As Variant) As Boolean
    On Error GoTo FAIL
    If IsNull(v) Or IsEmpty(v) Then GoTo FAIL
    If VarType(v) = vbString Then
        If Trim$(UCase$(v)) = "NA" Or Trim$(v) = "" Then GoTo FAIL
    End If
    If Not IsDate(v) Then GoTo FAIL
    If CDate(v) = 0 Then GoTo FAIL
    
    IsDateValid = True
    Exit Function
FAIL:
    IsDateValid = False
End Function

Private Function DateToStr(ByVal d As Date) As String
    DateToStr = Format(d, "yyyy-mm-dd")
End Function

Private Function NumToJSON(ByVal n As Double) As String
    NumToJSON = Replace(CStr(n), ",", ".")
End Function

Private Function ResourceTypeName(ByVal t As Long) As String
    Select Case t
        Case pjResourceTypeWork: ResourceTypeName = "Work"
        Case pjResourceTypeMaterial: ResourceTypeName = "Material"
        Case Else: ResourceTypeName = "Cost"
    End Select
End Function

Private Function PrettyJSON(ByVal s As String, Optional indent As Integer = 2) As String
    Dim i As Long, ch As String * 1, out As String
    Dim inStringFlag As Boolean, esc As Boolean
    Dim lvl As Long, tabs As String

    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)

        If esc Then
            out = out & ch
            esc = False

        ElseIf ch = "\" Then
            esc = True
            out = out & ch

        ElseIf ch = """" Then
            ' Toggle string mode and emit the quote
            inStringFlag = Not inStringFlag
            out = out & ch

        ElseIf inStringFlag Then
            ' Inside a JSON string ? copy verbatim
            out = out & ch

        Else
            Select Case ch
                Case "{", "["
                    out = out & ch & vbCrLf
                    lvl = lvl + 1
                    tabs = String(lvl * indent, " ")
                    out = out & tabs

                Case "}", "]"
                    out = out & vbCrLf
                    lvl = lvl - 1
                    If lvl < 0 Then lvl = 0
                    tabs = String(lvl * indent, " ")
                    out = out & tabs & ch

                Case ","
                    out = out & ch & vbCrLf & String(lvl * indent, " ")

                Case ":"
                    out = out & ": "

                Case " ", vbCr, vbLf, vbTab
                    ' skip whitespace outside strings

                Case Else
                    ' Any other non-whitespace char outside strings
                    out = out & ch
            End Select
        End If
    Next i

    PrettyJSON = out
End Function

Private Sub SortVariantArray(ByRef arr As Variant)
    On Error Resume Next
    If Not IsArray(arr) Then Exit Sub
    Dim i As Long, j As Long
    Dim temp As Variant

    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If CStr(arr(j)) < CStr(arr(i)) Then
                temp = arr(i)
                arr(i) = arr(j)
                arr(j) = temp
            End If
        Next j
    Next i
End Sub

