

   Sub tickerSymbol():

    Dim ticker As String
    Dim start_date As Long
    Dim end_date As Long
    Dim price_start As Double
    Dim price_end As Double
    Dim price_change As Double
    Dim percent_change As Variant
    Dim stock_volume As Double
    Dim year As Long
    
    Dim d As Integer
    
    Dim max As Double
    Dim ticker_max As String
    Dim min As Double
    Dim ticker_min As String
    Dim total_vol As Double
    Dim ticker_vol As String
    
    
    'define first and last days of the year
    year = Left(Cells(2, 2), 4)
   
    start_date = year & "0101"
  
   
   
  
    d = 2

    
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
  
    For i = 2 To LastRow
        
    
        'find row with start_date
        If Cells(i, 2).Value = start_date Then
            
        
            'list ticker symbol
            Cells(d, 9).Value = Cells(i, 1).Value
            
        
            'list price at start of year
            price_start = Cells(i, 3).Value
           
            
            'initiate summing stock_volume
            stock_volume = Cells(i, 7).Value
           
         
        'nested loop to determine row with end_date
        For j = i To LastRow
            
           
            'find row with end_date for ticker symbol
           If Cells(j + 1, 1).Value <> Cells(j, 1) Then
           
                Cells(j, 2).Value = end_date
                    
                'determine price_end
                price_end = Cells(j, 6).Value
               ' Cells(d, 15).Value = Cells(j, 6).Value
            
                'calculate and list price difference
                 Cells(d, 10).Value = price_end - price_start
            
                'calculate and list percent change, "n/a" if price_start = 0
                    If price_start > 0 Then
                    Cells(d, 11).Value = (price_end - price_start) / price_start
                    
                    Else
                        Cells(d, 11).Value = "n/a"
                    
                    End If
                 
                 'sum total stock_volume
                Cells(d, 12).Value = stock_volume + Cells(j, 7).Value
                   
                  
                'increment display row by 1
                d = d + 1
                
                Exit For
                
                Else
                    'add stock volume between start and end dates
                    stock_volume = stock_volume + Cells(j, 7).Value
            
            End If
            
         Next j

       
        End If
        
    Next i
    
    For c = 2 To LastRow
    
           ' Cells(c, 10).Font.ColorIndex = 1
            'Set cell color to red if negative change
            If Cells(c, 10) < 0 Then
            Cells(c, 10).Interior.ColorIndex = 3
            
            'Set cell color to green if positive change
            ElseIf Cells(c, 10) > 0 Then
            Cells(c, 10).Interior.ColorIndex = 4
            
            End If
            
    Next c
    
     
    max = Cells(2, 11).Value
    min = Cells(2, 11).Value
    total_vol = Cells(2, 12).Value
    
    LastRowPC = Cells(Rows.Count, 11).End(xlUp).Row
        For m = 3 To LastRowPC
            
            If IsNumeric(Cells(m, 11).Value) Then
                If Cells(m, 11).Value > max Then
                    max = Cells(m, 11).Value
                    ticker_max = Cells(m, 9).Value
                
                End If
        
            Cells(2, 16).Value = max
            Cells(2, 15).Value = ticker_max
  
        
        
                If Cells(m, 11).Value < min Then
                    min = Cells(m, 11).Value
                    ticker_min = Cells(m, 9).Value
                End If
        
            Cells(3, 16).Value = min
            Cells(3, 15).Value = ticker_min
            
    
                If Cells(m, 12).Value > total_vol Then
                    total_vol = Cells(m, 12).Value
                    ticker_vol = Cells(m, 9).Value
                End If
        
            Cells(4, 16).Value = total_vol
            Cells(4, 15).Value = ticker_vol
            
            End If
    
        Next m
    
End Sub




