Attribute VB_Name = "mod_GenericFunctions"
Option Explicit

Sub SetMacroMode(mode As Boolean)
    If mode Then
        Application.Calculation = xlCalculationManual
        Application.ScreenUpdating = False
        Application.EnableEvents = False
    Else
        Application.Calculation = xlCalculationAutomatic
        Application.ScreenUpdating = True
        Application.EnableEvents = True
        Application.StatusBar = False
    End If
End Sub

Sub RelockSheet(ws As Worksheet)
'ensure macro is allowed to edit worksheet as UserInterfaceOnly resets to False after workbook is closed/open
    ws.Protect Password:=PW, userinterfaceonly:=True, AllowFiltering:=True
    ws.EnableOutlining = True
End Sub

Sub UnfilterSheet(ws As Worksheet)
    On Error Resume Next
    ws.AutoFilterMode = False
End Sub

Sub ClearData(dataHeaders As Range)

Dim rowsCount As Long
rowsCount = GetLastRow(dataHeaders) - dataHeaders.Row
rowsCount = Application.WorksheetFunction.Max(rowsCount, 1)
dataHeaders.Offset(1).Resize(rowsCount).Clear

End Sub

Function GetLastRow(rng As Range, Optional columnsCount As Long) As Long

    Dim curWs As Worksheet
    Dim iRow As Long
    Dim iCol As Long
    Dim curLastRow As Long
    Dim finalLastRow As Long
    Dim curCellValue As String
    
    Set curWs = rng.Worksheet
    If columnsCount < 1 Then columnsCount = rng.Columns.Count
    
    For iCol = 1 To columnsCount
        iRow = curWs.Cells(1048576, rng.Column).Offset(0, iCol - 1).End(xlUp).Row
        curCellValue = "temp"
        
        On Error Resume Next
            curCellValue = curWs.Cells(iRow, rng.Column).Offset(0, iCol - 1).Value
        On Error GoTo 0
        
        Do Until curCellValue <> ""
            iRow = iRow - 1
            
            If iRow < 1 Then GoTo next_iCol
            On Error Resume Next
                curCellValue = curWs.Cells(iRow, rng.Column).Offset(0, iCol - 1).Value
            On Error GoTo 0
        Loop
        
        curLastRow = iRow
        finalLastRow = WorksheetFunction.Max(curLastRow, finalLastRow)
    
next_iCol:
    Next iCol
    
    finalLastRow = WorksheetFunction.Max(finalLastRow, 1)
    GetLastRow = finalLastRow

End Function

Function SmartUnion(curRange As Range, addRange As Range) As Range
    If addRange Is Nothing Then
        Set SmartUnion = curRange
    Else
        If curRange Is Nothing Then
            Set SmartUnion = addRange
        Else
            Set SmartUnion = Union(curRange, addRange)
        End If
    End If
End Function

