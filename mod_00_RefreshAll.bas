Attribute VB_Name = "mod_00_RefreshAll"
Option Explicit

Sub Button_RefreshAll()

    If Not DEBUG_MODE Then On Error GoTo closeMacro
    mod_GenericFunctions.SetMacroMode True

    Call mod_01_RefreshData.RefreshData
    Call mod_02_Workings_Hierarchy.RefreshWorkingsHierarchy
    Call mod_02_Workings_Lists.RefreshWorkingsList
    
closeMacro:
    mod_GenericFunctions.SetMacroMode False
    If Err Then
        MsgBox "An error occurred! Macro terminated!", vbCritical
    Else
        MsgBox "Completed successfully!", vbInformation
    End If


End Sub


