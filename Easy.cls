VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Easy()
'Iterate over all worksheets:
For Each ws In Worksheets
'Label headers:
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Total Stock Volume"
'Assign variable types:
    Dim Max As Long
    Dim num As Long
    Dim Ticker As String
    Dim Volume As Long
'Set to zero:
    Ticker = Blank
'num will be the total number of unique stocks
    num = 0
'Calculate max to equal maximum rows in sheet:
    Max = ws.Cells(Rows.Count, 1).End(xlUp).Row
'Iterate over rows:
    For i = 2 To Max
'For each new stock name, increase num by 1 and make Ticker equal
'the name of the stock and Volume equal the volume:
        If ws.Cells(i, 1).Value <> Ticker Then
            num = num + 1
            Volume = ws.Cells(i, 7).Value
            Ticker = ws.Cells(i, 1).Value
            ws.Cells(num + 1, 9).Value = Ticker
            ws.Cells(num + 1, 10).Value = Volume
'For each repeated stock name, add the additional volume:
        Else
            ws.Cells(num + 1, 10).Value = ws.Cells(num + 1, 10).Value + ws.Cells(i, 7).Value
        End If
        Next i
Next ws
End Sub
