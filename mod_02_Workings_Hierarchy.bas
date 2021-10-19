Attribute VB_Name = "mod_02_Workings_Hierarchy"
Option Explicit

Private rng_workingsHierarchyHeaders As Range
Private ws_source As Worksheet

'NO BUTTON IN USE
'Sub Button_RefreshSummaryHierarchy()
'
'    If Not DEBUG_MODE Then On Error GoTo closeMacro
'    mod_GenericFunctions.SetMacroMode True
'
'    Call RefreshWorkingsHierarchy
'
'closeMacro:
'    SetMacroMode False
'    If Err Then
'        MsgBox "An error occurred! Macro terminated!", vbCritical
'    Else
'        MsgBox "Completed successfully!", vbInformation
'    End If
'End Sub

Sub RefreshWorkingsHierarchy()
    
    'module level variables
    Set ws_source = WS_DATASOURCE
    Set rng_workingsHierarchyHeaders = WS_WORKINGS.Range("HierarchyHeaders")
    
    mod_GenericFunctions.RelockSheet ws_source
    mod_GenericFunctions.UnfilterSheet ws_source
    mod_GenericFunctions.RelockSheet WS_WORKINGS
    mod_GenericFunctions.UnfilterSheet WS_WORKINGS

    mod_GenericFunctions.ClearData rng_workingsHierarchyHeaders
    
    Call copyRawDataToWorkings
    Call formatWorkingsData

End Sub

Private Sub copyRawDataToWorkings()

    Dim colsCount As Integer
    Dim sourceColPosition() As Long
    Dim sourceDataLastRow As Long
    Dim curHeaderValue As String
    Dim sourceDataRng As Range
    Dim iCol As Integer
    
    colsCount = rng_workingsHierarchyHeaders.Columns.Count
    ReDim sourceColPosition(1 To colsCount)
    
    sourceDataLastRow = mod_GenericFunctions.GetLastRow(ws_source.Cells(HEADERROW_PT, 1), 20)
    
    For iCol = 1 To colsCount
        curHeaderValue = rng_workingsHierarchyHeaders.Resize(1, 1).Cells(1, iCol).Value
        sourceColPosition(iCol) = WorksheetFunction.Match(curHeaderValue, ws_source.Rows(HEADERROW_PT), 0)
        Set sourceDataRng = ws_source.Range(ws_source.Cells(HEADERROW_PT, sourceColPosition(iCol)), ws_source.Cells(sourceDataLastRow, sourceColPosition(iCol)))
    
        sourceDataRng.Copy
        rng_workingsHierarchyHeaders.Resize(, 1).Cells(, iCol).PasteSpecial xlPasteValues
        Application.CutCopyMode = False
    Next iCol

End Sub


Private Sub formatWorkingsData()
    
    Dim rowOffset As Long
    Dim rowsCount As Long
    Dim rowEmpty  As Long
    Dim colOffset   As Integer
    Dim colsCount   As Integer
    Dim colsArray() As Variant
    
    rowsCount = mod_GenericFunctions.GetLastRow(rng_workingsHierarchyHeaders) - rng_workingsHierarchyHeaders.Row
    colsCount = rng_workingsHierarchyHeaders.Columns.Count
    
    ReDim colsArray(0 To colsCount - 1)
    For colOffset = 0 To colsCount - 1
        colsArray(colOffset) = colOffset + 1
    Next colOffset
    
    'rows needs to be cleared, as it is faster to .removeDuplicates (empty rows) than .delete multiple rows, even in one instruction
    For rowOffset = 1 To rowsCount
        For colOffset = 0 To colsCount - 1
            If rng_workingsHierarchyHeaders.Resize(, 1).Offset(rowOffset, colOffset).Value = "" Then
                rng_workingsHierarchyHeaders.Offset(rowOffset).Clear
                GoTo Next_rowOffset
            End If
        Next colOffset
Next_rowOffset:
    Next rowOffset
    
    rng_workingsHierarchyHeaders.Offset(1).Resize(rowsCount).RemoveDuplicates Columns:=(colsArray), Header:=xlNo
    
    'find and remove empty row
    rowOffset = rng_workingsHierarchyHeaders.Resize(, 1).End(xlDown).Row - rng_workingsHierarchyHeaders.Row + 1
    rng_workingsHierarchyHeaders.Offset(rowOffset).Delete xlUp
    
    rowsCount = mod_GenericFunctions.GetLastRow(rng_workingsHierarchyHeaders)
    
    With WS_WORKINGS.Sort
        .SortFields.Clear
        For colOffset = 0 To colsCount - 1
            .SortFields.Add Key:=rng_workingsHierarchyHeaders.Resize(rowsCount, 1).Offset(, colOffset) ', SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        Next colOffset
        .SetRange rng_workingsHierarchyHeaders.Resize(rowsCount + 1)
        .Header = xlYes
        .Apply
    End With

End Sub


