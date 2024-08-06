Attribute VB_Name = "Module4"
Sub ClearSheet()
ÊÊÊ Dim ws As Worksheet
ÊÊÊ Dim lastRow As Long
ÊÊÊ Dim lastCol As Long
ÊÊÊ
ÊÊÊÊSet ws = ThisWorkbook.Sheets("BOM + ITEM")
ÊÊÊ lastRow = ws.UsedRange.Rows.Count
ÊÊÊ lastCol = ws.UsedRange.Columns.Count
ÊÊÊ
ÊÊÊÊws.Range(ws.Cells(4, 1), ws.Cells(lastRow, lastCol)).ClearContents
ÊÊÊÊws.Range(ws.Cells(4, 1), ws.Cells(lastRow, lastCol)).Interior.Color = RGB(255, 255, 255)
ÊÊÊ
End Sub
Sub ClearSheetButton()
ÊÊÊ Call ClearSheet
End Sub
