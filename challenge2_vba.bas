Attribute VB_Name = "Module1"
Sub challenge2():

For Each ws In Worksheets

    ws.Range("K1").Value = "Ticker"
    ws.Range("L1").Value = "Yearly Change"
    ws.Range("M1").Value = "Percent Change"
    ws.Range("N1").Value = "Total Stock Volume"
    ws.Range("R2").Value = "Greatest % Increase"
    ws.Range("R3").Value = "Greatest % Decrase"
    ws.Range("R4").Value = "Greatest Total Volume"
    ws.Range("S1").Value = "Ticker"
    ws.Range("T1").Value = "Value"
    

    Dim tckr As String
    Dim yrly_change As Double
    Dim close_lastrow As Double
    Dim opening_firstrow As Double
    Dim prct_change As Double
    Dim tot_stck_vol As Integer
    Dim sol_row As Integer
    
 
    
    opening_firstrow = ws.Cells(2, 3).Value
    tot_stck_vol = 0
    sol_row = 2

    
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    
    For i = 2 To LastRow
            
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1) Then
        
            tckr = ws.Cells(i, 1).Value
            ws.Range("K" & sol_row).Value = tckr
            

            tot_stock_vol = tot_stock_vol + ws.Cells(i, 7).Value
            ws.Range("N" & sol_row).Value = tot_stock_vol
            

            close_lastrow = ws.Cells(i, 6).Value
            yrly_change = close_lastrow - opening_firstrow
            ws.Range("L" & sol_row) = yrly_change
            prct_change = (yrly_change / opening_firstrow)
            ws.Range("M" & sol_row).Value = prct_change
            
        
            opening_firstrow = ws.Cells(i + 1, 3)

            sol_row = sol_row + 1
            
            tot_stock_vol = 0
           
        Else
            
            tot_stock_vol = tot_stock_vol + ws.Cells(i, 7).Value
              
        End If
        
    
    Next i
    
    ws.Range("M:M").NumberFormat = "0.00%"
    ws.Range("T2, T3").NumberFormat = "0.00%"
     
    second_lastrow = ws.Cells(Rows.Count, 12).End(xlUp).Row
    
    For i = 2 To second_lastrow
        
        If ws.Cells(i, 12).Value > 0 Then
            ws.Cells(i, 12).Interior.ColorIndex = 4
            
        ElseIf ws.Cells(i, 12).Value < 0 Then
            ws.Cells(i, 12).Interior.ColorIndex = 3
        
        Else
            ws.Cells(i, 12).Interior.ColorIndex = 5
        
        End If
        
    Next i
    
    
    Dim inc_tckr, dec_tcker, vol_tckr As String
    Dim inc_val, dec_val, vol_val As Double
    Dim counter As Double
    
    third_lastrow = ws.Cells(Rows.Count, 13).End(xlUp).Row
    
    
    inc_val = ws.Cells(2, 13).Value
    
    For i = 2 To third_lastrow
        
        
        If ws.Cells(i, 13).Value > inc_val Then
            inc_val = ws.Cells(i, 13).Value
            ws.Cells(2, 20).Value = inc_val
            inc_tckr = ws.Cells(i, 11).Value
            ws.Cells(2, 19).Value = inc_tckr
            
        End If
             
    Next i

    
    dec_val = ws.Cells(2, 13).Value

    For i = 2 To third_lastrow
    
        If ws.Cells(i, 13).Value < dec_val Then
            dec_val = ws.Cells(i, 13).Value
            ws.Cells(3, 20).Value = dec_val
            dec_tckr = ws.Cells(i, 11).Value
            ws.Cells(3, 19).Value = dec_tckr
            
        End If
    
    Next i
     
     
    vol_val = ws.Cells(2, 14).Value

    For i = 2 To third_lastrow
    
        If ws.Cells(i, 14).Value > vol_val Then
            vol_val = ws.Cells(i, 14).Value
            ws.Cells(4, 20).Value = vol_val
            vol_tckr = ws.Cells(i, 11).Value
            ws.Cells(4, 19).Value = vol_tckr
            
        End If
    
    Next i

    
    ws.Range("R:R, S:S, T:T, K1, L1, M1, N1").EntireColumn.AutoFit
    ws.Range("L:L").NumberFormat = "0.00"
   

Next ws


End Sub




