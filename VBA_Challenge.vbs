Sub Stock()

'Define variables for worksheet
Dim Ticker As String
Dim StockVolume As Double
StockVolume = 0
Dim Open_Price As Double

'Set open price
Open_Price = Cells(2, 3).Value

Dim Close_Price As Double
Dim Yearly_Change As Double
Dim Percent_Change As Double
Dim Summary_Ticker_Row  As Integer
Summary_Ticker_Row = 2


'Set header names
Cells(1, 9).Value = "TickerName"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"

'Count number of rows in Column 1
LastRow = Cells(Rows.Count, 1).End(xlUp).Row
'MsgBox (Lastrow)

'Find duplicates in the tickers
For i = 2 To LastRow
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

'Set TickerName
TickerName = Cells(i, 1).Value
'MsgBox (TickerName)

'Set Stock Volume
StockVolume = StockVolume + Cells(i, 7).Value

'Add TickerName in Summary total
Range("I" & Summary_Ticker_Row).Value = TickerName

'Add Stock Volume in Summary total
Range("L" & Summary_Ticker_Row).Value = StockVolume

'Set close price
Close_Price = Cells(i, 6).Value

'Find Yearly Change, take the close price and subtract by open price
Yearly_Change = (Close_Price - Open_Price)

'Yearly Change into Summary total
Range("J" & Summary_Ticker_Row).Value = Yearly_Change

If (Open_Price = 0) Then
    Percent_Change = 0
    
    Else
        Percent_Change = (Yearly_Change / Open_Price)
    
    
End If

'Yearly change for each ticker in summary
Range("K" & Summary_Ticker_Row).Value = Percent_Change
Range("K" & Summary_Ticker_Row).NumberFormat = "0.00%"

'Reset the Row Counter
Summary_Ticker_Row = Summary_Ticker_Row + 1

'Reset Stock Volume
StockVolume = 0

'Reset the Opening Price
Open_Price = Cells(i + 1, 3)

Else

    'Add Stock Volume
    StockVolume = StockVolume + Cells(i, 7).Value


End If


Next i

'Find Last Row of Summary
Lastrow_Summary = Cells(Rows.Count, 9).End(xlUp).Row

'Highlight cells - Positive change to green, negative change to red
For i = 2 To Lastrow_Summary
        If Cells(i, 10).Value < 0 Then
                Cells(i, 10).Interior.ColorIndex = 3
        Else
                Cells(i, 10).Interior.ColorIndex = 4
        End If
        
Next i

'Define Variables for greatest increase section
Dim GreatestIncrease As Double
GreatestIncrease = 0
Dim GreatestDecrease As Double
GreatestDecrease = 0
Dim GreatestVolume As Double
GreatestVolume = 0

'Set header title for greatest % increase section

Cells(2, 15).Value = "Greatest % Increase"
Cells(3, 15).Value = "Greatest % Decrease"
Cells(4, 15).Value = "Greatest Total Volume"
Cells(1, 16).Value = "TickerName"
Cells(1, 17).Value = "Value"


'Loop for final results
For i = 2 To LastRow

If Range("K" & i).Value > Range("Q2").Value Then
      Range("Q2").Value = Range("K" & i).Value
      Range("P2").Value = Range("I" & i).Value
    
End If

 If Range("K" & i).Value < Range("Q3").Value Then
        Range("Q3").Value = Range("K" & i).Value
        Range("P3").Value = Range("I" & i).Value
        
End If

 If Range("L" & i).Value > Range("Q4").Value Then
        Range("Q4").Value = Range("L" & i).Value
        Range("P4").Value = Range("I" & i).Value

End If

    Next i

'Add % Symbol
Range("Q2").NumberFormat = "0.00%"
Range("Q3").NumberFormat = "0.00%"

End Sub


