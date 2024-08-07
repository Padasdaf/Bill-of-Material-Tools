Attribute VB_Name = "Module3"
Sub DescriptionModify()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long, j As Long
    Dim primaryRow As Long
    Dim currentString As String
    Dim firstD As Boolean
    Dim endPos As Long

    Set ws = ThisWorkbook.Sheets("BOM + Item")
    lastRow = ws.UsedRange.Rows.Count
    For i = 1 To lastRow
        If ws.Cells(i, 12).Value = "EA (each)" Then
            primaryRow = i
            currentString = ws.Cells(primaryRow, 10).Value
            firstD = True
            endPos = Len(currentString)

            If ws.Cells(primaryRow, 6).Value = "M" Then
                If currentString <> "" Then currentString = currentString & ";"
                currentString = currentString & ws.Cells(primaryRow, 7).Value
            End If

            For j = i + 1 To lastRow
                If ws.Cells(j, 12).Value = "EA (each)" Then
                    Exit For
                End If
                If ws.Cells(j, 6).Value = "M" Then
                    If currentString <> "" Then currentString = currentString & ";"
                    currentString = currentString & ws.Cells(j, 7).Value
                End If
            Next j

            If ws.Cells(primaryRow, 6).Value = "D" Then
                If currentString <> "" Then currentString = currentString & ";"
                If firstD Then currentString = currentString & "DWG:"
                If firstD Then firstD = False
                currentString = currentString & ws.Cells(primaryRow, 8).Value
                If ws.Cells(primaryRow, 17).Value <> "" Then ws.Cells(primaryRow, 17).Value = ws.Cells(primaryRow, 17).Value & ";"
                ws.Cells(primaryRow, 17).Value = ws.Cells(primaryRow, 17).Value & ws.Cells(primaryRow, 8).Value
            End If
            For j = i + 1 To lastRow
                If ws.Cells(j, 12).Value = "EA (each)" Then
                    Exit For
                End If
                If ws.Cells(j, 6).Value = "D" Then
                    If currentString <> "" Then currentString = currentString & ";"
                    If firstD Then currentString = currentString & "DWG:"
                    If firstD Then firstD = False
                    currentString = currentString & ws.Cells(j, 8).Value
                    If ws.Cells(primaryRow, 17).Value <> "" Then ws.Cells(primaryRow, 17).Value = ws.Cells(primaryRow, 17).Value & ";"
                    ws.Cells(primaryRow, 17).Value = ws.Cells(primaryRow, 17).Value & ws.Cells(j, 8).Value
                End If
            Next j
            ws.Cells(primaryRow, 10).Value = currentString
            ws.Cells(primaryRow, 10).Characters(0, Len(currentString)).Font.Color = RGB(255, 0, 0)
            ws.Cells(primaryRow, 10).Characters(0, endPos).Font.Color = RGB(0, 0, 0)
            ws.Cells(primaryRow, 17).Characters(0, Len(ws.Cells(primaryRow, 17).Value)).Font.Color = RGB(255, 0, 0)
        End If
    Next i
End Sub
Sub DescriptionModifyButton()
    Call DescriptionModify
End Sub
