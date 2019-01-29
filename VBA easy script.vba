Sub Wallstreetloop()

    Dim ws As Worksheet
    For Each ws In Worksheets
    ws.Activate
    Dim ticker As String
    
    Dim Total_Stock_Volume As Double
    Total_Stock_Volume = 0
    
    Dim summary_table_row As Integer
    summary_table_row = 2
    
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Total Stock Volume"
    
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To lastrow
    
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
            ticker = Cells(i, 1).Value
            
            Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
            
            ws.Range("I" & summary_table_row).Value = ticker
            ws.Range("J" & summary_table_row).Value = Total_Stock_Volume
            
            summary_table_row = summary_table_row + 1
            
            Total_Stock_Volume = 0
            
        Else
            
            Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
            
        End If
        
    Next i
    
Next ws

End Sub

