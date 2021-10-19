Attribute VB_Name = "mod_02_Workings_Lists"
Option Explicit
Private srcHeaders As Range
Private destHeaders As Range

Sub RefreshWorkingsList()

Dim colOffset As Integer
Dim colsCount As Integer
Dim rowsCount As Long

Set srcHeaders = WS_WORKINGS.Range("HierarchyHeaders")
Set destHeaders = WS_WORKINGS.Range("UniqueListHeaders")

mod_GenericFunctions.ClearData destHeaders
colsCount = srcHeaders.Columns.Count
rowsCount = mod_GenericFunctions.GetLastRow(srcHeaders) - srcHeaders.Row

For colOffset = 0 To colsCount - 1
    srcHeaders.Offset(1, colOffset).Resize(rowsCount, 1).Copy
    destHeaders.Offset(1, colOffset).Resize(, 1).PasteSpecial xlPasteValues
    Application.CutCopyMode = False
    destHeaders.Offset(, colOffset).Resize(rowsCount + 1, 1).RemoveDuplicates Columns:=Array(1), Header:=xlYes
Next colOffset

End Sub
