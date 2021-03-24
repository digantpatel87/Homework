
Sub Main()
    Dim MaxRow As Double
    
    MaxRow = Cells(Rows.Count, "A").End(xlUp).Row
    
    
    'Populate distinct Tickers
    Dim NextTickerCell As Double
    NextTickerCell = 2
    Dim RunningTicker As String
    Dim RunningDistTicker As String
    Dim RunningOpen As Double
    Dim RunningClose As Double
    Dim RunningVolumn As Double
    Dim MaxnumberOfGivenTicker As Double
    Dim PreviousMaxnumberOfGivenTicker As Double
    
    PreviousMaxnumberOfGivenTicker = 0
    
    Cells(1, "I").Value = "Ticker"
    Cells(1, "J").Value = "Yearly Change"
    Cells(1, "K").Value = "Percentage Change"
    Cells(1, "L").Value = "Total Sock Volume"
    
    Cells(2, "O").Value = "Greatest % increase"
    Cells(3, "O").Value = "Greatest % decrease"
    Cells(4, "O").Value = "Greatest total volume"
    Cells(1, "P").Value = "Ticker"
    Cells(1, "Q").Value = "Value"
    
    'Loop for each row
    For i = 2 To MaxRow
        RunningTicker = Cells(i, "A")
           
        'Check if ticker exists in I column
        If Application.WorksheetFunction.CountIf(Range("I2:I" & MaxRow), RunningTicker) = 0 Then
        
            'Set new ticker in I column
            Cells(NextTickerCell, "I").Value = RunningTicker
                                 
            'Get Open of given ticker
            RunningOpen = Cells(i, 3).Value
            
            'get max records of given ticker
            MaxnumberOfGivenTicker = Application.WorksheetFunction.CountIf(Range("A2:A" & MaxRow), RunningTicker)
            
            'Add max given ticker records to running number to use it to get closing number
            RunningClose = Cells(1 + MaxnumberOfGivenTicker + PreviousMaxnumberOfGivenTicker, "F").Value
            
            'Get Difference
            Cells(NextTickerCell, "J").Value = RunningClose - RunningOpen
            
            If RunningOpen <> 0 Then
                'Calculate percentage change
                Cells(NextTickerCell, "K").Value = (RunningClose - RunningOpen) / RunningOpen
                Cells(NextTickerCell, "K").NumberFormat = "0.00%"
            Else
                Cells(NextTickerCell, "K").Value = 0
                Cells(NextTickerCell, "K").NumberFormat = "0.00%"
            End If
            
            
            'Get the sum of given Ticker using Sumif
            Cells(NextTickerCell, "L").Value = Application.WorksheetFunction.SumIf(Range("A2:A" & MaxRow), RunningTicker, Range("G2:G" & MaxRow))
            
            'Set this to be used for getting close number
            PreviousMaxnumberOfGivenTicker = MaxnumberOfGivenTicker + PreviousMaxnumberOfGivenTicker
            'Set this to determine next cell for I column
            NextTickerCell = NextTickerCell + 1
                                  
                               
        End If
        
              
    Next i
          
    'Conditional formatting
    'clear any existing conditional formatting
    Range("J2", Range("J2").End(xlDown)).FormatConditions.Delete

    'Conditional Formating for Positive value in Green color
    With Range("J2", Range("J2").End(xlDown)).FormatConditions.Add(xlCellValue, xlGreater, "=0")
        .Interior.Color = vbGreen
    End With
    
    'Conditional Formating for Negative value in red color
    With Range("J2", Range("J2").End(xlDown)).FormatConditions.Add(xlCellValue, xlLess, "=0")
        .Interior.Color = vbRed
    End With
    
    
    'Bonus
    Dim BonusMaxRow As Long
    Dim BonusRunningTicker As String
    Dim BonusRunningPercentage As Double
    Dim BonusRunningTotalStockVolumn As Double
    Dim BonusTotalStockVolumn As Double
           
        
        
    BonusMaxRow = Cells(Rows.Count, "I").End(xlUp).Row
    
    'Range("L1:L" & BonusMaxRow).NumberFormat = "0"
    
    Cells(2, "Q").Value = 0
    Cells(3, "Q").Value = 0
    Cells(4, "Q").Value = 0
    
    Cells(2, "Q").NumberFormat = "0.00%"
    Cells(3, "Q").NumberFormat = "0.00%"
    
    
    For j = 2 To BonusMaxRow
        BonusRunningTicker = Cells(j, "I")
        BonusRunningPercentage = Cells(j, "K")
        BonusRunningTotalStockVolumn = Cells(j, "L")
        
        If BonusRunningPercentage > Cells(2, "Q").Value Then
            Cells(2, "Q").Value = BonusRunningPercentage
            Cells(2, "P").Value = BonusRunningTicker
        End If
        
        If BonusRunningPercentage < Cells(3, "Q").Value Then
            Cells(3, "Q").Value = BonusRunningPercentage
            Cells(3, "P").Value = BonusRunningTicker
        End If
        
        
        BonusTotalStockVolumn = Cells(4, "Q").Value
        
        If BonusRunningTotalStockVolumn > BonusTotalStockVolumn Then
            Cells(4, "Q").Value = BonusRunningTotalStockVolumn
            Cells(4, "P").Value = BonusRunningTicker
        End If
        
    
    Next j
    
    
End Sub



