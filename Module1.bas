Attribute VB_Name = "Module1"

Sub StockMarketChecker()

Dim ws As Worksheet
For Each ws In Worksheets

'Set Table Column and Row Headers
ws.Cells(1, "I").Value = "Ticker"
ws.Cells(1, "J").Value = "Yearly Change"
ws.Cells(1, "K").Value = "Percent Change"
ws.Cells(1, "L").Value = "Total Stock Volume"

ws.Cells(1, "P").Value = "Ticker"
ws.Cells(1, "Q").Value = "Value"
ws.Cells(2, "O").Value = "Greatest % Increase"
ws.Cells(3, "O").Value = "Greatest % Decrease"
ws.Cells(4, "O").Value = "Greatest Total Volume"

'Set Variable Data Type
Dim Ticker As String
Dim Lastrow As Long
Dim i As Long
Dim j As Long

'Calculate last row number
Lastrow = ws.Cells(Rows.Count, "A").End(xlUp).Row

Dim OpenPrice As Double
Dim ClosePrice As Double
Dim YearlyChange As Double
Dim PriceChangePercent As Double

Dim StockRow As Long
StockRow = 2

j = 2

'Loop through rows
For i = 2 To Lastrow
    If ws.Cells(i + 1, "A").Value <> ws.Cells(i, "A").Value Then
        Ticker = ws.Cells(i, "A").Value
        ws.Cells(StockRow, "I").Value = Ticker

        OpenPrice = ws.Cells(j, "C").Value
        ClosePrice = ws.Cells(i, "F").Value
        YearlyChange = ClosePrice - OpenPrice
        ws.Cells(StockRow, "J").Value = YearlyChange
        
        'Set COnditional Formating for Yearly Change
            If YearlyChange < 0 Then
                ws.Cells(StockRow, "J").Interior.ColorIndex = 3
            ElseIf YearlyChange > 0 Then
                ws.Cells(StockRow, "J").Interior.ColorIndex = 4
            End If
        
        'Set Conditional Formating for Percent Change, and stop 0 error
            If OpenPrice <> 0 Then
                PriceChangePercent = ((ClosePrice - OpenPrice) / OpenPrice)
                ws.Cells(StockRow, "K").Value = PriceChangePercent
                ws.Cells(StockRow, "K").Style = "Percent"
            ElseIf OpenPrice = 0 Then
                ws.Cells(StockRow, "K").Value = 0
                ws.Cells(StockRow, "K").Style = "Percent"
            End If
        
        'Calculate Stock Volume
        ws.Cells(StockRow, "L").Value = WorksheetFunction.Sum(Range(ws.Cells(j, "G"), ws.Cells(i, "G")))
        
        'Set next row for Stock Row
        StockRow = StockRow + 1
        'Set new value to calculate next stock
        j = i + 1
    End If
Next i

Dim LastRow2 As Long
Dim IncreasePercent As Double
Dim DecreasePercent As Double
Dim TotalVolume As Double

'Set start values for calculations
IncreasePercent = ws.Cells(2, "K").Value
DecreasePercent = ws.Cells(2, "K").Value
TotalVolume = ws.Cells(2, "L").Value

'Calculate last row number for new summarised table
LastRow2 = ws.Cells(Rows.Count, "I").End(xlUp).Row

'Loop through rows for new summarised table
For i = 2 To LastRow2
    If ws.Cells(i, "K").Value > IncreasePercent Then
        IncreasePercent = ws.Cells(i, "K").Value
        ws.Cells(2, "P").Value = ws.Cells(i, "I").Value
        ws.Cells(2, "Q").Value = ws.Cells(i, "K").Value
    Else
        IncreasePercent = IncreasePercent
    
    If ws.Cells(i, "K").Value < DecreasePercent Then
        DecreasePercent = ws.Cells(i, "K").Value
        ws.Cells(3, "P").Value = ws.Cells(i, "I").Value
        ws.Cells(3, "Q").Value = ws.Cells(i, "K").Value
    Else
        DecreasePercent = DecreasePercent
        
    If ws.Cells(i, "L").Value > TotalVolume Then
        TotalVolume = ws.Cells(i, "L").Value
        ws.Cells(4, "P").Value = ws.Cells(i, "I").Value
        ws.Cells(4, "Q").Value = ws.Cells(i, "L").Value
    Else
        TotalVolume = TotalVolume
        
        
    End If
    End If
    End If
    
Next i
    
Next ws

End Sub
