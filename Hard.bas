Attribute VB_Name = "Module3"
Sub Hard()
'It is necessary to run the Moderate Module before running this code
'Iterate over all worksheets:
For Each ws In Worksheets
'Label headers:
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    ws.Cells(2, 15).Value = "Greatest Percent Increase"
    ws.Cells(3, 15).Value = "Greatest Percent Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
'Assign value types and set original values to zero:
    Dim Max As Long
    Dim GreatestPercentIncrease As Double
    Dim GreatestPercentDecrease As Double
    Dim GreatestVolume As Double
    GreatestPercentIncrease = 0
    GreatestPercentDecrease = 0
    GreatestVolume = 0
'Calculate Max as the total amount of rows
    Max = ws.Cells(Rows.Count, 9).End(xlUp).Row
'Iterate over the rows
       For i = 2 To Max
'Check if each percent increase is greater than the current greatest, overwrite the cell if it is:
            If ws.Cells(i, 11) > GreatestPercentIncrease Then
                GreatestPercentIncrease = ws.Cells(i, 11)
                ws.Cells(2, 16).Value = ws.Cells(i, 11).Value
                ws.Cells(2, 17).Value = ws.Cells(i, 9).Value
            End If
'Check if each percent increase is less than the current least, overwrite the cell if it is:
            If ws.Cells(i, 11) < GreatestPercentDecrease Then
                GreatestPercentDecrease = ws.Cells(i, 11)
                ws.Cells(3, 16).Value = ws.Cells(i, 11).Value
                ws.Cells(3, 17).Value = ws.Cells(i, 9).Value
            End If
'Check if each volume is greater than the current greatest volume, overwrite the cell if it is:
            If ws.Cells(i, 12) > GreatestVolume Then
                GreatestVolume = ws.Cells(i, 12)
                ws.Cells(4, 16).Value = ws.Cells(i, 12).Value
                ws.Cells(4, 17).Value = ws.Cells(i, 9).Value
            End If
        Next i
'Format the percentage values as percents
        ws.Cells(3, 16).Value = FormatPercent(ws.Cells(3, 16).Value)
        ws.Cells(2, 16).Value = FormatPercent(ws.Cells(2, 16).Value)
Next ws
    
End Sub
