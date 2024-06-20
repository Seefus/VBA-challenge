Sub Fun()
 Dim lastRow As Long
Dim ws As Worksheet
Dim sheetNames As Variant
 Dim ticker As String
 Dim row As Long
 Dim volume As Double
 Dim op As Double
 Dim cl As Double
 Dim QC As Double
 Dim i As Long
 Dim pc As Double
 'second part
Dim j As Long
Dim lastrow2 As Long
Dim git As String
Dim gi As Double
Dim di As Double
Dim dit As String
Dim vi As Double
Dim vit As String
sheetNames = Array("Q1", "Q2", "Q3", "Q4")
For Each sheetName In sheetNames

 Set ws = ThisWorkbook.Sheets(sheetName)
 lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
 lastrow2 = ws.Cells(ws.Rows.Count, "I").End(xlUp).row
 row = 2
 volume = 0
 op = ws.Cells(2, 3).Value
 For i = 2 To lastRow
 If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
 
 ticker = ws.Cells(i, 1).Value
 ws.Cells(row, 9).Value = ticker
 volume = volume + ws.Cells(i, 7).Value
 ws.Cells(row, 12).Value = volume
 cl = ws.Cells(i, 6).Value
 QC = cl - op
 ws.Cells(row, 10).Value = QC
 ws.Cells(row, 10).NumberFormat = "0.00"

 If QC > 0 Then
 ws.Cells(row, 10).Interior.ColorIndex = 4
 ElseIf QC < 0 Then
ws.Cells(row, 10).Interior.ColorIndex = 3
 Else
ws.Cells(row, 10).Interior.ColorIndex = xlNone
 End If
 
 pc = ((cl - op) / op)
 ws.Cells(row, 11).Value = pc
ws.Cells(row, 11).NumberFormat = "0.00%"
 op = ws.Cells(i + 1, 3).Value
 
 row = row + 1
 volume = 0
 
 Else
 volume = volume + ws.Cells(i, 7).Value
 
 
 
 End If
 Next i
  gi = -1E+307
  di = 1
  vi = 1
 For j = 2 To row - 1
        If ws.Cells(j, 11).Value > gi Then
            gi = ws.Cells(j, 11).Value
            git = ws.Cells(j, 9).Value
        End If
        If ws.Cells(j, 11).Value < di Then
        di = ws.Cells(j, 11).Value
        dit = ws.Cells(j, 9).Value
        End If
        If ws.Cells(j, 12).Value > vi Then
        vi = ws.Cells(j, 12).Value
        vit = ws.Cells(j, 9).Value
        End If
        
    Next j
 ws.Cells(2, 16).Value = git
 ws.Cells(2, 17).Value = gi
 ws.Cells(2, 17).NumberFormat = "0.00%"
 ws.Cells(3, 16).Value = dit
 ws.Cells(3, 17).Value = di
 ws.Cells(3, 17).NumberFormat = "0.00%"
 ws.Cells(4, 16).Value = vit
 ws.Cells(4, 17).Value = vi
 Next sheetName
 
End Sub
