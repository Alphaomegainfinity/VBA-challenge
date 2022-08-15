Sub stock_exchange()
Dim yearly_change As Double
Dim ticker_name As String
Dim ticker_change As Double
Dim stock_volume As Double
Dim last_row As Double
Dim year_open As Double
Dim year_close As Double
Dim Summary_row As Integer
Dim percent_change As Double
Dim ticker_max As String
Dim ticker_min As String
Dim ticker_total_value As String

'Find the last row in column Ticker
Dim last_row1 As Double
Dim max_value As Double
Dim min_value As Double
Dim max_total_value As Double

'Speed up VBA running
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

'Looping through each worksheet
For Each Worksheet In Worksheets
stock_volume = 0
Summary_row = 2

'set the first ticker open value
year_open = Worksheet.Cells(2, 3).Value

    'Checking the last row in each worksheet
    ActiveSheet.UsedRange
    last_row = ActiveSheet.UsedRange.Rows(ActiveSheet.UsedRange.Rows.Count).Row
    
    'Add the name Ticker, Yearly Change, Percent Change and Total Stock Volume to new columns
    Worksheet.Range("I1").Value = "Ticker"
    Worksheet.Range("J1").Value = "Yearly Change"
    Worksheet.Range("K1").Value = "Percent Change"
    Worksheet.Range("L1").Value = "Total Stock Volume"
    
        
    For i = 2 To last_row
        'Create a summary table for each stock
        If Worksheet.Cells(i + 1, 1) <> Worksheet.Cells(i, 1) Then
            'create a list of ticker name
            ticker_name = Worksheet.Cells(i, 1).Value
                        
            'set the total stock volume
            stock_volume = stock_volume + Worksheet.Cells(i, 7).Value
            
            'Finding the last value of a stock close and calculate the yearly change
            year_close = Worksheet.Cells(i, 6).Value
            yearly_change = year_close - year_open
            
            'Calculate the percentage change
            percent_change = (yearly_change / year_open)
                                 
            'Add value to summary table
            Worksheet.Range("I" & Summary_row).Value = ticker_name
            Worksheet.Range("L" & Summary_row).Value = stock_volume
            Worksheet.Range("J" & Summary_row).Value = yearly_change
            Worksheet.Range("K" & Summary_row).Value = percent_change
            Worksheet.Range("K" & Summary_row).NumberFormat = "0.00%"
            
                       
            'Adding color to yearly change column
            If yearly_change < 0 Then
                Worksheet.Range("J" & Summary_row).Interior.ColorIndex = 3
            Else
                Worksheet.Range("J" & Summary_row).Interior.ColorIndex = 4
            End If
            
            Summary_row = Summary_row + 1
            
            'Reset the open price for the next stock
            year_open = Worksheet.Cells(i + 1, 3).Value
            
            stock_volume = 0
        
        Else:
             'Worksheet.ticker_name = Worksheet.Cells(i, 1)
            stock_volume = stock_volume + Worksheet.Cells(i, 7).Value
        End If
             
        
    Next i

'Bonus

    'Add the name Ticker, Value, Greatest %Increase, Greatest %Decrease, Greatest Total Volume to new columns
    Worksheet.Range("P1").Value = "Ticker"
    Worksheet.Range("Q1").Value = "Value"
    Worksheet.Range("O2").Value = "Greatest % Increase"
    Worksheet.Range("O3").Value = "Greatest % Decrease"
    Worksheet.Range("O4").Value = "Greatest Total Volume"
    
    
        
    last_row1 = Worksheet.Cells(Rows.Count, 9).End(xlUp).Row
    
    max_value = 0
    min_value = 0
    max_total_volume = 0
    
        For r = 2 To (last_row1 - 1)
    
        'Find the Max value in Percent Change column
            If Worksheet.Cells(r, 11) >= max_value Then
    
                max_value = Worksheet.Cells(r, 11).Value
                ticker_max = Worksheet.Cells(r, 9).Value
           
            End If
            
         ' Find the Min Value in Percent Change column
            If Worksheet.Cells(r, 11) <= min_value Then
    
                min_value = Worksheet.Cells(r, 11).Value
                ticker_min = Worksheet.Cells(r, 9).Value
           
            End If
            
          ' Find the greatest total volume from total stock volume column
            If Worksheet.Cells(r, 12) >= max_total_value Then
    
                max_total_value = Worksheet.Cells(r, 12).Value
                ticker_total_value = Worksheet.Cells(r, 9).Value
             End If
        Next r
        
        Worksheet.Range("Q2").Value = max_value
        Worksheet.Range("Q2").NumberFormat = "0.00%"
        Worksheet.Range("P2").Value = ticker_max
        
        Worksheet.Range("Q3").Value = min_value
        Worksheet.Range("Q3").NumberFormat = "0.00%"
        Worksheet.Range("P3").Value = ticker_min
        
        Worksheet.Range("Q4").Value = max_total_value
        Worksheet.Range("P4").Value = ticker_total_value
    
    
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic

Next Worksheet





End Sub

