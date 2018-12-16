Sub stock_market()
        
    Dim ws As Worksheet
    Dim Ticker As String
    Dim total_volume As Double
    Dim sum_table As Integer
    Dim Year_Change As Double
    Dim Percent_Change As Double
    
    

    For Each ws In Worksheets
    ws.Activate
    
    total_volume = 0
    
    sum_table = 2
    
    
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Total Volume"
    ws.Range("K1").Value = "Year Change"
    ws.Range("L1").Value = "Percent Change"
    
    Last = ActiveSheet.Range("B" & Rows.Count).End(xlUp).Row
    
        For i = 2 To Last
        
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            
                
                Ticker = Cells(i, 1).Value
                
                Year_Change = Cells(i, 6).Value - Cells(i, 3).Value
                
                Percent_Change = ((Cells(i, 6).Value - Cells(i, 3).Value) / Cells(i, 3).Value)
                
    
                
                total_volume = total_volume + Cells(i, 7).Value
                
                ws.Range("K" & sum_table).Value = Year_Change
                
                ws.Range("L" & sum_table).Value = Percent_Change
                
                ws.Range("I" & sum_table).Value = Ticker
                
                
                ws.Range("J" & sum_table).Value = total_volume
                
                
                sum_table = sum_table + 1
                
            
                total_volume = 0
            
           Else
            
                total_volume = total_volume + Cells(i, 7).Value
                
            End If
        
        Next i
        
    Next ws
    
End Sub