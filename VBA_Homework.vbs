Sub Stockdata()
  
    Dim total_volume As Double
    Dim ticker As String
    Dim ticker_counter As Integer
    
    For Each ws In Worksheets
    
        total_volume = 0
        ticker_counter = 2
        
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Total Stock Volume"
        
        For i = 2 To ws.Cells(Rows.Count, 1).End(xlUp).Row
            total_volume = total_volume + ws.Cells(i, 7).Value
            ticker = ws.Cells(i, 1).Value
            
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ws.Cells(ticker_counter, 9).Value = ticker
                ws.Cells(ticker_counter, 10).Value = total_volume
                total_volume = 0
                ticker_counter = ticker_counter + 1
            End If
        Next i

        ws.Columns("J").AutoFit
    
    Next ws

End Sub

Sub FirstStockdata()
    Columns("I:J").ClearContents
    Columns("I:J").ClearFormats
    Columns("I:J").UseStandardWidth = True

End Sub

Sub SecondStockdata()
    For Each ws In Worksheets
        ws.Columns("I:J").ClearContents
        ws.Columns("I:J").ClearFormats
        ws.Columns("I:J").UseStandardWidth = True
    Next ws
End Sub
    
Sub YearStock()

    Dim total_volume As Double
    Dim ticker As String
    Dim ticker_counter, ticker_open_close_counter As Double
    Dim yearly_open, yearly_end As Double
    
        total_volume = 0
        ticker_counter = 2
        ticker_open_close_counter = 2
        
        Range("I1").Value = "Ticker"
        Range("J1").Value = "Yearly Change"
        Range("K1").Value = "Percent Change"
        Range("L1").Value = "Total Stock Volume"
    
     For i = 2 To Cells(Rows.Count, 1).End(xlUp).Row

        total_vol = total_vol + Cells(i, 7).Value
        ticker = Cells(i, 1).Value
        yearly_open = Cells(ticker_open_close_counter, 3)
        
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            yearly_end = Cells(i, 6)
            Cells(ticker_counter, 9).Value = ticker
            Cells(ticker_counter, 10).Value = yearly_end - yearly_open
            
            If yearly_open = 0 Then
                Cells(ticker_counter, 11).Value = Null
            Else
                Cells(ticker_counter, 11).Value = (yearly_end - yearly_open) / yearly_open
            End If
            Cells(ticker_counter, 12).Value = total_vol
            
            If Cells(ticker_counter, 10).Value > 0 Then
                Cells(ticker_counter, 10).Interior.ColorIndex = 4
            Else
                Cells(ticker_counter, 10).Interior.ColorIndex = 3
            End If
            
            Cells(ticker_counter, 11).NumberFormat = "0.00%"
            
            total_volume = 0
            ticker_counter = ticker_counter + 1
            ticker_open_close_counter = i + 1
            
        End If
        
    Next i
    
    
    End Sub
  
Sub YearStocktwo()
    Dim total_vol As Double
    Dim ticker As String
    Dim ticker_counter, ticker_open_close_counter As Double
    Dim yearly_open, yearly_end As Double
    
    For Each ws In Worksheets
        total_volume = 0
        ticker_counter = 2
        ticker_open_close_counter = 2
        
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        
      LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    For i = 2 To LastRow


            total_volume = total_volume + ws.Cells(i, 7).Value
            ticker = ws.Cells(i, 1).Value
            yearly_open = ws.Cells(ticker_open_close_counter, 3)
             If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            yearly_end = ws.Cells(i, 6)
            ws.Cells(ticker_counter, 9).Value = ticker
            ws.Cells(ticker_counter, 10).Value = yearly_end - yearly_open
            If yearly_open = 0 Then
                ws.Cells(ticker_counter, 11).Value = Null
            Else
                ws.Cells(ticker_counter, 11).Value = (yearly_end - yearly_open) / yearly_open
            End If
            ws.Cells(ticker_counter, 12).Value = total_volume
            
            'Color the cell green if > 0, red if < 0
            If ws.Cells(ticker_counter, 10).Value > 0 Then
                ws.Cells(ticker_counter, 10).Interior.ColorIndex = 4
            Else
                ws.Cells(ticker_counter, 10).Interior.ColorIndex = 3
            End If
            
            ws.Cells(ticker_counter, 11).NumberFormat = "0.00%"
            
            total_volume = 0
            ticker_counter = ticker_counter + 1
            ticker_open_close_counter = i + 1
            
            End If
           
            
            Next i
            
            Next ws
            
        End Sub
           
        Sub Stockfour()
    Dim total_vol As Double
    Dim ticker As String
    Dim ticker_counter, ticker_open_close_counter As Double
    Dim yearly_open, yearly_end As Double
    
    For Each ws In Worksheets
        total_vol = 0
        ticker_counter = 2
        ticker_open_close_counter = 2
        
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        For i = 2 To ws.Cells(Rows.Count, 1).End(xlUp).Row
            total_vol = total_vol + ws.Cells(i, 7).Value
            ticker = ws.Cells(i, 1).Value
            yearly_open = ws.Cells(ticker_open_close_counter, 3)
            
            
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                yearly_end = ws.Cells(i, 6)
                ws.Cells(ticker_counter, 9).Value = ticker
                ws.Cells(ticker_counter, 10).Value = yearly_end - yearly_open
                
                If yearly_open = 0 Then
                    ws.Cells(ticker_counter, 11).Value = Null
                Else
                    ws.Cells(ticker_counter, 11).Value = (yearly_end - yearly_open) / yearly_open
                End If
                ws.Cells(ticker_counter, 12).Value = total_vol
                
                ' Color the cell green if > 0, red if < 0
                If ws.Cells(ticker_counter, 10).Value > 0 Then
                    ws.Cells(ticker_counter, 10).Interior.ColorIndex = 4
                Else
                    ws.Cells(ticker_counter, 10).Interior.ColorIndex = 3
                End If
                
                ws.Cells(ticker_counter, 11).NumberFormat = "0.00%"
                
                total_vol = 0
                ticker_counter = ticker_counter + 1
                ticker_open_close_counter = i + 1
            End If
            
        Next i

        ws.Columns("J").AutoFit
        ws.Columns("K").AutoFit
        ws.Columns("L").AutoFit

    Next ws
End Sub
Sub YearStockthree()
    Columns("I:L").ClearContents
    Columns("I:L").ClearFormats
    Columns("I:L").UseStandardWidth = True
End Sub

Sub Stocks()
    For Each ws In Worksheets
        ws.Columns("I:L").ClearContents
        ws.Columns("I:L").ClearFormats
        ws.Columns("I:L").UseStandardWidth = True
    Next ws
End Sub
        
        
        
    
    
        
    
    
    
    
    
    
    
    

    
    
    
