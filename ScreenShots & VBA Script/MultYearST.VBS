Sub MultYearST()

Application.ScreenUpdating = False

    'define everything
Dim ws As Worksheet
Dim ticker As String
Dim vol As Double
Dim year_open As Double
Dim year_close As Double
Dim yearly_change As Double
Dim percent_change As Double
Dim Summary_Table_Row As Integer

vol = 0

'run through each worksheet
For Each ws In ThisWorkbook.Worksheets
    'set headers
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
 'setup integers for loop
    Summary_Table_Row = 2
    
'loop
        For i = 2 To ws.UsedRange.Rows.Count
             If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
            'find all the values
                ticker = ws.Cells(i, 1).Value
                vol = vol + ws.Cells(i, 7).Value
                year_open = ws.Cells(i, 3).Value
                year_close = ws.Cells(i, 6).Value
                yearly_change = year_close - year_open
                    If yearly_change > 0 Then
                    ws.Range("J" & Summary_Table_Row).Interior.Color = vbGreen
                    Else
                    ws.Range("J" & Summary_Table_Row).Interior.Color = vbRed
                    End If
                    
                percent_change = (year_close - year_open) / year_close

            'insert values into summary
                ws.Cells(Summary_Table_Row, 9).Value = ticker
                ws.Cells(Summary_Table_Row, 10).Value = yearly_change
                ws.Cells(Summary_Table_Row, 11).Value = percent_change
                ws.Cells(Summary_Table_Row, 12).Value = vol
                Summary_Table_Row = Summary_Table_Row + 1

                vol = 0
             
             Else
                vol = vol + ws.Cells(i, 7).Value
            End If
             
    'finish loop
    Next i
        
        ws.Columns("K").NumberFormat = "0.00%"
Next ws

Application.ScreenUpdating = True

End Sub



