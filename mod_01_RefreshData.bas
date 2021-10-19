Attribute VB_Name = "mod_01_RefreshData"
Option Explicit

Sub Button_RefreshData()
    
    If Not DEBUG_MODE Then On Error GoTo closeMacro
    mod_GenericFunctions.SetMacroMode True

    Call RefreshData
    
closeMacro:
    SetMacroMode False
    If Err Then
        MsgBox "An error occurred! Macro terminated!", vbCritical
    Else
        MsgBox "Completed successfully!", vbInformation
    End If

End Sub

Sub RefreshData()

'here is a piece of code that refreshes Pivot Tables like the one in WS_DATASOURCE

End Sub
