Attribute VB_Name = "Module5"
Sub LevelHierarchyFiller()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim currentNumber As Variant

    Set ws = ThisWorkbook.Sheets("BOM + Item")
    lastRow = ws.UsedRange.Rows.Count
    currentNumber = Empty

    For i = 4 To lastRow
        If IsNumeric(ws.Cells(i, 5).Value) And ws.Cells(i, 5).Value <> "" Then
            currentNumber = ws.Cells(i, 5).Value
        ElseIf IsEmpty(ws.Cells(i, 5).Value) And Not IsEmpty(currentNumber) Then
            ws.Cells(i, 5).Value = currentNumber + 1
        End If
    Next i

End Sub
Sub LevelHierarchyFillerButton()
    Call LevelHierarchyFiller
End Sub
