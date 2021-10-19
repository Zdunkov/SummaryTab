Attribute VB_Name = "mod_03_SummaryHierarchy"
Option Explicit

Sub Button_RefreshSummaryHierarchy()

    If Not DEBUG_MODE Then On Error GoTo closeMacro
    mod_GenericFunctions.SetMacroMode True

    Call RefreshSummaryHierarchy
    
closeMacro:
    SetMacroMode False
    If Err Then
        MsgBox "An error occurred! Macro terminated!", vbCritical
    Else
        'MsgBox "Completed successfully!", vbInformation
    End If


End Sub

Sub RefreshSummaryHierarchy()

    mod_GenericFunctions.RelockSheet WS_SUMMARY
    
    WS_SUMMARY.Range("HierarchyHeaders").Resize(10000).Rows.Hidden = False

    populateSummaryHierarchy
    copyFormulas
    hideData
    
    WS_SUMMARY.Range("HierarchyHeaders").Offset.Rows.Hidden = True
    
End Sub

Private Sub populateSummaryHierarchy()
    
    Dim src_rowOffset As Long
    Dim src_headers As Range
    Dim dest_rowOffset As Long
    Dim dest_headers As Range
    Dim rowsCount As Long
    Dim colOffset As Integer
    Dim colsCount As Integer
    Dim cellValue As String
    Dim i As Integer
    Dim priorCol As Integer
    Dim lastColumnValue() As String
    

    Set src_headers = WS_WORKINGS.Range("HierarchyHeaders")
    Set dest_headers = WS_SUMMARY.Range("HierarchyHeaders")
    
    rowsCount = mod_GenericFunctions.GetLastRow(src_headers) - src_headers.Row
    dest_rowOffset = 1
    colsCount = src_headers.Columns.Count
    ReDim lastColumnValue(0 To colsCount - 1)
    
    mod_GenericFunctions.ClearData dest_headers
    
    
    For src_rowOffset = 1 To rowsCount
        For colOffset = 0 To colsCount - 1
            
            cellValue = src_headers.Resize(1, 1).Offset(src_rowOffset, colOffset).Value
            
            If cellValue = lastColumnValue(colOffset) Then GoTo Next_colOffset
            
            dest_headers.Resize(1, 1).Offset(dest_rowOffset, colOffset).Value = cellValue
            dest_headers.Resize(1, 1).Offset(dest_rowOffset, colOffset).Font.Color = RGB(0, 0, 0)
            
            For priorCol = 1 To colOffset
                dest_headers.Resize(1, 1).Offset(dest_rowOffset, priorCol - 1).Value = lastColumnValue(priorCol - 1)
                dest_headers.Resize(1, 1).Offset(dest_rowOffset, priorCol - 1).Font.Color = RGB(222, 222, 222)
            Next priorCol
            
            lastColumnValue(colOffset) = cellValue
            
            For i = colOffset + 1 To colsCount - 1
                lastColumnValue(i) = ""
            Next i
            
            dest_rowOffset = dest_rowOffset + 1
    
Next_colOffset:
        Next colOffset
    Next src_rowOffset

End Sub

Private Sub copyFormulas()
    
    Dim hierarchyHeadersRange As Range
    Dim rowsCount As Long
    
    Dim sourceRange As Range
    Dim destRange As Range
    
    Set hierarchyHeadersRange = WS_SUMMARY.Range("HierarchyHeaders")
    rowsCount = mod_GenericFunctions.GetLastRow(hierarchyHeadersRange) - hierarchyHeadersRange.Row
    
    Set sourceRange = WS_SUMMARY.Range("HierarchyFormulas")
    Set destRange = sourceRange.Offset(1).Resize(rowsCount)
    
    mod_GenericFunctions.ClearData sourceRange
    sourceRange.Copy destRange

End Sub


Private Sub hideData()

    Dim hierarchyHeadersRange As Range
    Dim columnsCount As Integer
    Dim IsHierarchyVisible() As Boolean
    Dim colOffset As Integer
    Dim rowOffset As Long
    Dim rowsCount As Long
    Dim currentRowLevel As Integer
    Dim rowsToHide As Range
    
    Set hierarchyHeadersRange = WS_SUMMARY.Range("HierarchyHeaders")
    columnsCount = hierarchyHeadersRange.Columns.Count
    rowsCount = mod_GenericFunctions.GetLastRow(hierarchyHeadersRange) - hierarchyHeadersRange.Row
    ReDim IsHierarchyVisible(1 To columnsCount)
    
    hierarchyHeadersRange.Offset(1).Resize(rowsCount).Rows.Hidden = False
    
    For colOffset = 1 To columnsCount
        IsHierarchyVisible(colOffset) = WS_SUMMARY.Shapes("CB_Hierarchy" & colOffset).ControlFormat.Value = 1
    Next colOffset
    
    For rowOffset = 1 To rowsCount
        For colOffset = 1 To columnsCount

            If hierarchyHeadersRange.Resize(1, 1).Offset(rowOffset, colOffset).Value <> "" Then GoTo Next_colOffset
                
            currentRowLevel = colOffset
            
            If Not IsHierarchyVisible(currentRowLevel) Then
                Set rowsToHide = mod_GenericFunctions.SmartUnion(rowsToHide, hierarchyHeadersRange.Offset(rowOffset, colOffset).EntireRow)
            End If
            
            GoTo Next_rowOffset

Next_colOffset:
        Next colOffset
Next_rowOffset:
    Next rowOffset

If Not rowsToHide Is Nothing Then
    rowsToHide.Rows.Hidden = True
End If

End Sub
