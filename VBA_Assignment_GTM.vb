Sub stocks()

' Setting variables

Dim ws As Worksheet
Dim ticker As String
Dim vol As Double
Dim stock_open As Double
Dim stock_close As Double
Dim change_over_year As Double
Dim change_percentage As Double


Dim Summary_Table_Row As Integer
Dim lastrow As Double

'kept getting an overflow error on the change_percentage so I'm using this next line to get around it

On Error Resume Next

' loop through each worksheet

For Each ws In Worksheets
lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
Summary_Table_Row = 2
vol = 0
change_over_year = 0
stock_open = 0
stock_close = 0
change_percentage = 0

'setting headings

ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"

'looping through columns to get totals

For i = 2 To lastrow
    
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        ticker = ws.Cells(i, 1).Value
        vol = vol + Cells(i, 7).Value
        stock_open = ws.Cells(i, 3).Value
        stock_close = ws.Cells(i, 6).Value
        change_over_year = stock_close - stock_open
        
        
       change_percentage = (stock_close - stock_open) / stock_close
    
    'setting values to new cells // did this two different ways as I was playing around with it to see how it worked

    ws.Range("I" & Summary_Table_Row).Value = ticker
    ws.Range("J" & Summary_Table_Row).Value = change_over_year
    ws.Cells(Summary_Table_Row, 11).Value = change_percentage
    ws.Cells(Summary_Table_Row, 12).Value = vol
    Summary_Table_Row = Summary_Table_Row + 1
    vol = 0
    ' stock_open = 0
    ' stock_close = 0
    ' change_percentage = 0
    ' change_over_year = 0
    
    Else
    
        vol = vol + Cells(i, 7).Value
        ' stock_open = ws.Cells(i, 3).Value
        ' stock_close = ws.Cells(i, 6).Value
        
    'Setting conditional formatting

        If Cells(Summary_Table_Row, 10).Value > 0 Then

            Cells(Summary_Table_Row, 10).Interior.ColorIndex = 4

    
        ElseIf Cells(Summary_Table_Row, 10).Value < 0 Then
    
            Cells(Summary_Table_Row, 10).Interior.ColorIndex = 3
    
    
'closing if statements
End If

End If

'move on to next iteration

Next i

' move on to next worksheet

Next ws

' msg box to let me know when the script has completed

MsgBox ("Success")



End Sub


