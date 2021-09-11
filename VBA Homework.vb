Sub WallStreet():

    For Each ws In Worksheets
        Dim Ticker As String
        Dim Opening_Value As Double
        Dim Closing_Value As Double
        Dim Yearly_Change As Double
        Dim Percent_Change As Double
        Dim Stock_Volume As Double
        Dim Summary_Table_Row As Long
        Dim Previous As Long
            
        Stock_Volume = 0
        Summary_Table_Row = 2
        Previous = 2

        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent. Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        For i = 2 To LastRow

            Stock_Volume = Stock_Volume + ws.Cells(i, 7).Value

            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                Ticker = ws.Cells(i, 1).Value
                Opening_Value = ws.Range("C" & Previous).Value
                Closing_Value = ws.Cells(i, 6).Value
                Yearly_Change = Closing_Value - Opening_Value

                If Opening_Value = 0 Then
                    Percent_Change = 0
                Else
                    Percent_Change = Yearly_Change / Opening_Value
                End If

                ws.Range("I" & Summary_Table_Row).Value = Ticker
                ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
                ws.Range("K" & Summary_Table_Row).Value = Percent_Change
                ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
                ws.Range("L" & Summary_Table_Row).Value = Stock_Volume
                
                Stock_Volume = 0
                
                If ws.Range("J" & Summary_Table_Row).Value >= 0 Then
                    ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                Else
                    ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                End If
            
                Summary_Table_Row = Summary_Table_Row + 1
                Previous = i + 1

            End If
                
        Next i

    Next ws
    
End Sub