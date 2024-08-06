Attribute VB_Name = "Module3"
Sub DescriptionModify()
ÊÊÊ Dim ws As Worksheet
ÊÊÊ Dim lastRow As Long
ÊÊÊ Dim i As Long, j As Long
ÊÊÊ Dim primaryRow As Long
ÊÊÊ Dim currentString As String
ÊÊÊ Dim firstD As Boolean
ÊÊÊ
ÊÊÊÊSet ws = ThisWorkbook.Sheets("BOM + Item")
ÊÊÊ lastRow = ws.UsedRange.Rows.Count
ÊÊÊ For i = 1 To lastRow
ÊÊÊÊÊÊÊ If ws.Cells(i, 12).Value = "EA (each)" Then
ÊÊÊÊÊÊÊÊÊÊÊ primaryRow = i
ÊÊÊÊÊÊÊÊÊÊÊ currentString = ws.Cells(primaryRow, 10).Value
ÊÊÊÊÊÊÊÊÊÊÊ firstD = True

ÊÊÊÊÊÊÊÊÊÊÊÊIf ws.Cells(primaryRow, 6).Value = "M" Then
ÊÊÊÊÊÊÊÊÊÊÊÊÊÊÊ If currentString <> "" Then currentString = currentString & ";"
ÊÊÊÊÊÊÊÊÊÊÊÊÊÊÊ currentString = currentString & ws.Cells(primaryRow, 7).Value
ÊÊÊÊÊÊÊÊÊÊÊ End If

ÊÊÊÊÊÊÊÊÊÊÊÊFor j = i + 1 To lastRow
ÊÊÊÊÊÊÊÊÊÊÊÊÊÊÊ If ws.Cells(j, 12).Value = "EA (each)" Then
ÊÊÊÊÊÊÊÊÊÊÊÊÊÊÊÊÊÊÊ Exit For
ÊÊÊÊÊÊÊÊÊÊÊÊÊÊÊ End If
ÊÊÊÊÊÊÊÊÊÊÊÊÊÊÊ If ws.Cells(j, 6).Value = "M" Then
ÊÊÊÊÊÊÊÊÊÊÊÊÊÊÊÊÊÊÊ If currentString <> "" Then currentString = currentString & ";"
ÊÊÊÊÊÊÊÊÊÊÊÊÊÊÊÊÊÊÊ currentString = currentString & ws.Cells(j, 7).Value
ÊÊÊÊÊÊÊÊÊÊÊÊÊÊÊ End If
ÊÊÊÊÊÊÊÊÊÊÊ Next j

ÊÊÊÊÊÊÊÊÊÊÊÊIf ws.Cells(primaryRow, 6).Value = "D" Then
ÊÊÊÊÊÊÊÊÊÊÊÊÊÊÊ If currentString <> "" Then currentString = currentString & ";"
ÊÊÊÊÊÊÊÊÊÊÊÊÊÊÊ If firstD Then currentString = currentString & "DWG:"
ÊÊÊÊÊÊÊÊÊÊÊÊÊÊÊ If firstD Then firstD = False
ÊÊÊÊÊÊÊÊÊÊÊÊÊÊÊ currentString = currentString & ws.Cells(primaryRow, 8).Value
ÊÊÊÊÊÊÊÊÊÊÊÊÊÊÊ If ws.Cells(primaryRow, 17).Value <> "" Then ws.Cells(primaryRow, 17).Value = ws.Cells(primaryRow, 17).Value & ";"
ÊÊÊÊÊÊÊÊÊÊÊÊÊÊÊ ws.Cells(primaryRow, 17).Value = ws.Cells(primaryRow, 17).Value & ws.Cells(primaryRow, 8).Value
ÊÊÊÊÊÊÊÊÊÊÊ End If
ÊÊÊÊÊÊÊÊÊÊÊ For j = i + 1 To lastRow
ÊÊÊÊÊÊÊÊÊÊÊÊÊÊÊ If ws.Cells(j, 12).Value = "EA (each)" Then
ÊÊÊÊÊÊÊÊÊÊÊÊÊÊÊÊÊÊÊ Exit For
ÊÊÊÊÊÊÊÊÊÊÊÊÊÊÊ End If
ÊÊÊÊÊÊÊÊÊÊÊÊÊÊÊ If ws.Cells(j, 6).Value = "D" Then
ÊÊÊÊÊÊÊÊÊÊÊÊÊÊÊÊÊÊÊ If currentString <> "" Then currentString = currentString & ";"
ÊÊÊÊÊÊÊÊÊÊÊÊÊÊÊÊÊÊÊ If firstD Then currentString = currentString & "DWG:"
ÊÊÊÊÊÊÊÊÊÊÊÊÊÊÊÊÊÊÊ If firstD Then firstD = False
ÊÊÊÊÊÊÊÊÊÊÊÊÊÊÊÊÊÊÊ currentString = currentString & ws.Cells(j, 8).Value
ÊÊÊÊÊÊÊÊÊÊÊÊÊÊÊÊÊÊÊ If ws.Cells(primaryRow, 17).Value <> "" Then ws.Cells(primaryRow, 17).Value = ws.Cells(primaryRow, 17).Value & ";"
ÊÊÊÊÊÊÊÊÊÊÊÊÊÊÊÊÊÊÊ ws.Cells(primaryRow, 17).Value = ws.Cells(primaryRow, 17).Value & ws.Cells(j, 8).Value
ÊÊÊÊÊÊÊÊÊÊÊÊÊÊÊ End If
ÊÊÊÊÊÊÊÊÊÊÊ Next j
ÊÊÊÊÊÊÊÊÊÊÊ ws.Cells(primaryRow, 10).Value = currentString
ÊÊÊÊÊÊÊ End If
ÊÊÊ Next i
End Sub
Sub DescriptionModifyButton()
ÊÊÊ Call DescriptionModify
End Sub
