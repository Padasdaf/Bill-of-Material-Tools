Attribute VB_Name = "Module3"
Sub DescriptionModify()
ΚΚΚ Dim ws As Worksheet
ΚΚΚ Dim lastRow As Long
ΚΚΚ Dim i As Long, j As Long
ΚΚΚ Dim primaryRow As Long
ΚΚΚ Dim currentString As String
ΚΚΚ Dim firstD As Boolean
ΚΚΚ
ΚΚΚΚSet ws = ThisWorkbook.Sheets("BOM + Item")
ΚΚΚ lastRow = ws.UsedRange.Rows.Count
ΚΚΚ For i = 1 To lastRow
ΚΚΚΚΚΚΚ If ws.Cells(i, 12).Value = "EA (each)" Then
ΚΚΚΚΚΚΚΚΚΚΚ primaryRow = i
ΚΚΚΚΚΚΚΚΚΚΚ currentString = ws.Cells(primaryRow, 10).Value
ΚΚΚΚΚΚΚΚΚΚΚ firstD = True

ΚΚΚΚΚΚΚΚΚΚΚΚIf ws.Cells(primaryRow, 6).Value = "M" Then
ΚΚΚΚΚΚΚΚΚΚΚΚΚΚΚ If currentString <> "" Then currentString = currentString & ";"
ΚΚΚΚΚΚΚΚΚΚΚΚΚΚΚ currentString = currentString & ws.Cells(primaryRow, 7).Value
ΚΚΚΚΚΚΚΚΚΚΚ End If

ΚΚΚΚΚΚΚΚΚΚΚΚFor j = i + 1 To lastRow
ΚΚΚΚΚΚΚΚΚΚΚΚΚΚΚ If ws.Cells(j, 12).Value = "EA (each)" Then
ΚΚΚΚΚΚΚΚΚΚΚΚΚΚΚΚΚΚΚ Exit For
ΚΚΚΚΚΚΚΚΚΚΚΚΚΚΚ End If
ΚΚΚΚΚΚΚΚΚΚΚΚΚΚΚ If ws.Cells(j, 6).Value = "M" Then
ΚΚΚΚΚΚΚΚΚΚΚΚΚΚΚΚΚΚΚ If currentString <> "" Then currentString = currentString & ";"
ΚΚΚΚΚΚΚΚΚΚΚΚΚΚΚΚΚΚΚ currentString = currentString & ws.Cells(j, 7).Value
ΚΚΚΚΚΚΚΚΚΚΚΚΚΚΚ End If
ΚΚΚΚΚΚΚΚΚΚΚ Next j

ΚΚΚΚΚΚΚΚΚΚΚΚIf ws.Cells(primaryRow, 6).Value = "D" Then
ΚΚΚΚΚΚΚΚΚΚΚΚΚΚΚ If currentString <> "" Then currentString = currentString & ";"
ΚΚΚΚΚΚΚΚΚΚΚΚΚΚΚ If firstD Then currentString = currentString & "DWG:"
ΚΚΚΚΚΚΚΚΚΚΚΚΚΚΚ If firstD Then firstD = False
ΚΚΚΚΚΚΚΚΚΚΚΚΚΚΚ currentString = currentString & ws.Cells(primaryRow, 8).Value
ΚΚΚΚΚΚΚΚΚΚΚΚΚΚΚ If ws.Cells(primaryRow, 17).Value <> "" Then ws.Cells(primaryRow, 17).Value = ws.Cells(primaryRow, 17).Value & ";"
ΚΚΚΚΚΚΚΚΚΚΚΚΚΚΚ ws.Cells(primaryRow, 17).Value = ws.Cells(primaryRow, 17).Value & ws.Cells(primaryRow, 8).Value
ΚΚΚΚΚΚΚΚΚΚΚ End If
ΚΚΚΚΚΚΚΚΚΚΚ For j = i + 1 To lastRow
ΚΚΚΚΚΚΚΚΚΚΚΚΚΚΚ If ws.Cells(j, 12).Value = "EA (each)" Then
ΚΚΚΚΚΚΚΚΚΚΚΚΚΚΚΚΚΚΚ Exit For
ΚΚΚΚΚΚΚΚΚΚΚΚΚΚΚ End If
ΚΚΚΚΚΚΚΚΚΚΚΚΚΚΚ If ws.Cells(j, 6).Value = "D" Then
ΚΚΚΚΚΚΚΚΚΚΚΚΚΚΚΚΚΚΚ If currentString <> "" Then currentString = currentString & ";"
ΚΚΚΚΚΚΚΚΚΚΚΚΚΚΚΚΚΚΚ If firstD Then currentString = currentString & "DWG:"
ΚΚΚΚΚΚΚΚΚΚΚΚΚΚΚΚΚΚΚ If firstD Then firstD = False
ΚΚΚΚΚΚΚΚΚΚΚΚΚΚΚΚΚΚΚ currentString = currentString & ws.Cells(j, 8).Value
ΚΚΚΚΚΚΚΚΚΚΚΚΚΚΚΚΚΚΚ If ws.Cells(primaryRow, 17).Value <> "" Then ws.Cells(primaryRow, 17).Value = ws.Cells(primaryRow, 17).Value & ";"
ΚΚΚΚΚΚΚΚΚΚΚΚΚΚΚΚΚΚΚ ws.Cells(primaryRow, 17).Value = ws.Cells(primaryRow, 17).Value & ws.Cells(primaryRow, 8).Value
ΚΚΚΚΚΚΚΚΚΚΚΚΚΚΚ End If
ΚΚΚΚΚΚΚΚΚΚΚ Next j
ΚΚΚΚΚΚΚΚΚΚΚ ws.Cells(primaryRow, 10).Value = currentString
ΚΚΚΚΚΚΚ End If
ΚΚΚ Next i
End Sub
Sub DescriptionModifyButton()
ΚΚΚ Call DescriptionModify
End Sub
