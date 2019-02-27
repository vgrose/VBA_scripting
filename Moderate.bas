Attribute VB_Name = "Module1"
Sub Moderate()
'Iterate over all worksheets:
For Each ws In Worksheets
'Label headers:
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
'Assign value types:
    Dim Max As Long
    Dim num As Long
    Dim Ticker As String
    Dim Volume As Long
    Dim OpenPrice As Double
    Dim ClosePrice As Double
    Dim YearlyChange As Double
    Dim PercentChange As Double
'Set the current ticker to be blank and the number of stocks to be zero:
    Ticker = Blank
    num = 0
'Calculate "Max" to be the total amount of rows:
    Max = ws.Cells(Rows.Count, 1).End(xlUp).Row
'Iterate over all rows
    For i = 2 To Max
'If the stock value is new, increase the number of stocks, write its volume and stock name
'into cells, and set "OpenPrice" to be its opening price:
        If ws.Cells(i, 1).Value <> Ticker Then
            num = num + 1
            Volume = ws.Cells(i, 7).Value
            Ticker = ws.Cells(i, 1).Value
            OpenPrice = ws.Cells(i, 3).Value
            ws.Cells(num + 1, 9).Value = Ticker
            ws.Cells(num + 1, 12).Value = Volume
'If the stock value is not new...
        Else
'Add the current volume to the total volume
            ws.Cells(num + 1, 12).Value = ws.Cells(num + 1, 12).Value + ws.Cells(i, 7).Value
'If
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
'set the closing price, and calculate the yearly change:
                ClosePrice = ws.Cells(i, 3).Value
                YearlyChange = ClosePrice - OpenPrice
'If there are no zero values, calculate percent change:
                If YearlyChange <> 0 Then
                    If OpenPrice <> 0 Then
                        PercentChange = YearlyChange / OpenPrice
                        ws.Cells(num + 1, 10).Value = YearlyChange
                        ws.Cells(num + 1, 11).Value = FormatPercent(PercentChange)
'Color yearly change cells if positive or negative:
                            If YearlyChange > 0 Then
                                ws.Cells(num + 1, 10).Interior.ColorIndex = 4
                            Else
                                ws.Cells(num + 1, 10).Interior.ColorIndex = 3
                            End If
'if the opening price was zero, calculate percent change:
                    ElseIf OpenPrice = 0 Then
                        PercentChange = 1
                    End If
'if the percent change was zero, calculate percent change:
                Else
                        PercentChange = 0
                        ws.Cells(num + 1, 10).Value = YearlyChange
                        ws.Cells(num + 1, 11).Value = FormatPercent(PercentChange)
                        ws.Cells(num + 1, 10).Interior.ColorIndex = 4
                End If
            End If
        End If
        Next i
Next ws
End Sub

   
