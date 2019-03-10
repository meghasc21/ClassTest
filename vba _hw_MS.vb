Sub Stock()
'Variable Declarations

Dim ticker As String
Dim volume As Double
Dim i As Double, j As Double
Dim ws As Worksheet
Dim LR As Double
Dim openP As Double
Dim closeP As Double
Dim PERCH As Long


'For Every Sheet
For Each ws In Worksheets

'Last Row of Column A
LR = ws.Cells(Rows.Count, 1).End(xlUp).Row
LL = ws.Cells(Rows.Count, 12).End(xlUp).Row

'Updating Row Header
ws.Range("H1").Value = "Ticker"
ws.Range("I1").Value = "Total Stock Volume"

'Setting initial value
volume = 0
openP = ws.Cells(2, 3).Value
closeP = 0
j = 2

'For loop for the rows

For i = 2 To LR

'If loop to check against next row and update final columns

If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

ws.Cells(j, 8).Value = ws.Cells(i, 1).Value
volume = volume + ws.Cells(i, 7).Value
ws.Cells(j, 9).Value = volume

j = j + 1

End If
Next i
Next ws

End Sub
