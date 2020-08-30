# VBA-Homework

Sub Visual_Basic()
For Each ws In Worksheets

ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"
ws.Cells(2, 15).Value = "Greatest % Increase"
ws.Cells(3, 15).Value = "Greatest % Decrease"
ws.Cells(4, 15).Value = "Greatest Total Volume"
ws.Cells(1, 16).Value = "Ticker"
ws.Cells(1, 17).Value = "Value"

Dim Ticker_Name As String
Dim Stock_Total As Double
Stock_Total = 0

Dim Yearly_Open_Change As String
Dim Yearly_Closing_Change As String
Dim Yearly_Change As String
Dim Percent_Change As String
Dim Summary_Table_Row As Integer
Summary_Table_Row = 2


lastrow = Cells(Rows.Count, 1).End(xlUp).Row

  For i = 2 To lastrow

     If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        Ticker_Name = ws.Cells(i, 1).Value
        ws.Range("I" & Summary_Table_Row).Value = Ticker_Name

        Yearly_Open_Change = ws.Cells(i, 3).Value
        Yearly_Closing_Change = ws.Cells(i, 6).Value
        Yearly_Change = Yearly_Closing_Change - Yearly_Open_Change
        ws.Cells(Summary_Table_Row, 10) = Yearly_Change
        Percent_Change = Yearly_Closing_Change / Yearly_Open_Change
        ws.Cells(Summary_Table_Row, 11) = Percent_Change
        ws.Cells(Summary_Table_Row, 11).NumberFormat = "0.00%"
        
        Stock_Total = Stock_Total + ws.Cells(i, 7).Value
        ws.Range("L" & Summary_Table_Row).Value = Stock_Total
        Summary_Table_Row = Summary_Table_Row + 1
        Stock_Total = 0
        
     Else
        Stock_Total = Stock_Total + ws.Cells(i, 7).Value

     End If
         

     If ws.Cells(i, 10).Value >= 0 Then
        ws.Cells(i, 10).Interior.ColorIndex = 4

     ElseIf ws.Cells(i, 10).Value < 0 Then
        ws.Cells(i, 10).Interior.ColorIndex = 3
      
     End If
    
     If ws.Cells(i, 11).Value = ws.Application.WorksheetFunction.Max(Range("K:K")) Then
        ws.Cells(2, 17).Value = ws.Cells(i, 11)
        ws.Cells(2, 17).NumberFormat = "0.00%"
        ws.Cells(2, 16).Value = ws.Cells(I, 9).Value
     End If
    
     If ws.Cells(i, 11).Value = ws.Application.WorksheetFunction.Min(Range("K:K")) Then
        ws.Cells(3, 17).Value = ws.Cells(i, 11)
        ws.Cells(3, 17).NumberFormat = "0.00%"
        ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
     End If
     
     If ws.Cells(i, 12).Value = ws.Application.WorksheetFunction.Max(Range("L:L")) Then
        ws.Cells(4, 17).Value = ws.Cells(i, 12)
        ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
     End If
   
   Next i
   
Next ws

End Sub