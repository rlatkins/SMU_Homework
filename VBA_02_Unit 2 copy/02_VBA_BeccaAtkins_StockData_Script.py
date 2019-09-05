Sub Stock_Data():
' Create Column Labels
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"
' Set variable for holding the Ticker Rating
Dim Ticker_Rating As String
' Set variable for Stock Total
Dim Stock_Volume_Total As Double
Stock_Volume_Total = 0
' Set variable for Open Total
Dim Stock_Open_Total As Double
Stock_Open_Total = 0
' Set variable for Close Total
Dim Stock_Close_Total As Double
Stock_Close_Total = 0
' Set variable for Yearly Change
Dim Yearly_Change As Double
Yearly_Change = 0
' Set variable for Percent Change
Dim Percent_Change As Double
Percent_Change = 0
' Set variable for Last Row
Dim LastRow As Double
LastRow = Cells(Rows.Count, 1).End(xlUp).Row
' Keep track of the location of each stock in a summary table
Dim Summary_Table_Row As String
Summary_Table_Row = 2
' Loop though the Tickers
For i = 2 To LastRow
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        Ticker = Cells(i, 1).Value
        Stock_Volume_Total = Stock_Volume_Total + Cells(i, 7).Value
        Stock_Open_Total = Stock_Open_Total + Cells(i, 3).Value
        Stock_Close_Total = Stock_Close_Total + Cells(i, 6).Value
        Yearly_Change = Stock_Close_Total - Stock_Open_Total
        Percent_Change = (((Stock_Close_Total - Stock_Open_Total) / Stock_Open_Total) * 100)
        ' Print the Credit Card Brand in the Summary Table
        Range("I" & Summary_Table_Row).Value = Ticker
        Range("J" & Summary_Table_Row).Value = Yearly_Change
            If Yearly_Change >= 0 Then
            Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
            ElseIf Yearly_Change < 0 Then
            Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
    End If
        Range("K" & Summary_Table_Row).Value = Percent_Change
        Range("K" & Summary_Table_Row).Select
        Selection.Style = "Percent"
        Range("L" & Summary_Table_Row).Value = Stock_Volume_Total
        Summary_Table_Row = Summary_Table_Row + 1
        Stock_Volume_Total = 0
    Else
        Stock_Volume_Total = Stock_Volume_Total + Cells(i, 7).Value
    End If

    
Next i

End Sub
