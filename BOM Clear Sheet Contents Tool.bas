Attribute VB_Name = "Module4"
Sub ClearSheet()
��� Dim ws As Worksheet
��� Dim lastRow As Long
��� Dim lastCol As Long
���
����Set ws = ThisWorkbook.Sheets("BOM + ITEM")
��� lastRow = ws.UsedRange.Rows.Count
��� lastCol = ws.UsedRange.Columns.Count
���
����ws.Range(ws.Cells(4, 1), ws.Cells(lastRow, lastCol)).ClearContents
���
End Sub
Sub ClearSheetButton()
��� Call ClearSheet
End Sub
