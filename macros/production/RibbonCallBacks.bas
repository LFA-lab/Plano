Attribute VB_Name = "RibbonCallbacks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private gRibbon As Object
Private Const DEBUG_RIBBON As Boolean = True

' OnRibbonLoad can stay as-is; Project will pass an object. (Optional: make it parameterless too.)
Public Sub OnRibbonLoad(ByVal ribbon As Object)
    Set gRibbon = ribbon
    If DEBUG_RIBBON Then
        MsgBox "Ribbon loaded (Plano)."
    End If
    Debug.Print "Ribbon loaded (Plano)."
End Sub

' Project-safe onAction: accept control parameter
Public Sub GenerateDashboard(ByVal control As Object)
    MsgBox "GenerateDashboard invoked."
    RunImport
End Sub
