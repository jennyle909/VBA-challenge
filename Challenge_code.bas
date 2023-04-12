Attribute VB_Name = "Module1"
Sub stockchallenge():

'declaring variables
Dim Ticker As String
Dim Totalstockvol As Variant
Dim Summary_Table_Row As Integer
Dim closingprice As Double
Dim openingprice As Double
Dim Maxpercentchange As Double
Dim WorksheetName As String

For Each ws In Worksheets
Totalstockvol = 0
lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
Summary_Table_Row = 2
WorksheetName = ws.Name

'Naming cell columns
ws.Cells(1, 9) = "ticker"
ws.Cells(1, 10) = "Yearly Change"
ws.Cells(1, 11) = "Percent change"
ws.Cells(1, 12) = "Total stock volume"
ws.Cells(1, 16) = "ticker"
ws.Cells(1, 17) = "Value"
ws.Cells(2, 15) = "Greatest % Increase"
ws.Cells(3, 15) = "Greatest % Decrease"
ws.Cells(4, 15) = "Greatest total volume"

For i = 2 To lastrow
    'Finding Ticker and total stock volume
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
      Ticker = ws.Cells(i, 1).Value
      Totalstockvol = Totalstockvol + ws.Cells(i, 7).Value
      ws.Range("I" & Summary_Table_Row).Value = Ticker
      ws.Range("L" & Summary_Table_Row).Value = Totalstockvol
      
      'determining closing price
      closingprice = ws.Cells(i, 6).Value
      ws.Range("j" & Summary_Table_Row).Value = closingprice - openingprice
      ws.Cells(Summary_Table_Row, 11) = ((closingprice - openingprice) / openingprice) * 100
      Percentchange = ws.Cells(Summary_Table_Row, 11)
      
      Summary_Table_Row = Summary_Table_Row + 1
      Totalstockvol = 0
    Else
      Totalstockvol = Totalstockvol + ws.Cells(i, 7).Value
    End If
    
    'Check if current row is the first trading day of the year
    If ws.Cells(i, 1) <> Ticker And Right(ws.Cells(i, 2).Value, 4) = "0102" Then
        openingprice = ws.Cells(i, 3)
    End If
    
    'change cell colors for yearly change
    If ws.Cells(i, 10) > 0 Then
       ws.Cells(i, 10).Interior.ColorIndex = 4
    ElseIf ws.Cells(i, 10) < 0 Then
        ws.Cells(i, 10).Interior.ColorIndex = 3
    Else
        ws.Cells(i, 10).Interior.ColorIndex = 0
   End If
Next i

'greatest percent increase using the maxfunction
ws.Cells(2, 17) = WorksheetFunction.Max(ws.Range("k:k"))

'greatest percent decrease using the minfunction
ws.Cells(3, 17) = WorksheetFunction.Min(ws.Range("k:k"))

'greatest total volume using the maxfunction
ws.Cells(4, 17) = WorksheetFunction.Max(ws.Range("l:l"))

Next ws
End Sub
