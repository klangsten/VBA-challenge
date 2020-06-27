Attribute VB_Name = "Module1"
Sub StockMarketAnalysis():

    'Column Headers
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    
    'Variables
    Dim TickerCol As Integer
    TickerCol = 9
    
    Dim PrintRow As Integer
    PrintRow = 2
    
    Dim TotalVolumeCol As Integer
    TotalVolumeCol = 12
    
    Dim lastrow As Double
    lastrow = Cells(Rows.Count, "A").End(xlUp).Row
    
    Dim TotalVolume As Double
    TotalVolume = 0
    
    Dim RawVolumeCol As Integer
    RawVolumeCol = 7
    
    Dim OpenPrice As Long
    OpenPrice = Cells(2, 3)
    
    Dim ClosePrice As Double
    ClosePrice = 0
    
    Dim ChangeInPriceCol As Integer
    ChangeInPriceCol = 10
    
    Dim PctChangeInPrice As Double
    PctChangeInPrice = 0
    
    Dim ChangeInPrice As Double
    
    Dim PctChangeCol As Integer
    PctChangeCol = 11
    
    Dim Green As Integer
    Green = 4
    
    Dim Red As Integer
    Red = 3
    
    'Pick out unique ticker symbols

   For i = 2 To lastrow
   
        Dim CurrentCell As String
        CurrentCell = Cells(i, 1).Value
        
        Dim NextCell As String
        NextCell = Cells(i + 1, 1).Value
        
        If CurrentCell = NextCell Then
            
            'Add Volume for each identical ticker symbol
            TotalVolume = TotalVolume + Cells(i, RawVolumeCol).Value
        
        
        Else
            'Add current row
            TotalVolume = TotalVolume + Cells(i, RawVolumeCol).Value
            
            'Print Total Volume
            Cells(PrintRow, TotalVolumeCol).Value = TotalVolume
            
            'Set Total Volume to Zero
            TotalVolume = 0
            
            'Print unique ticker symbol
            Cells(PrintRow, TickerCol).Value = CurrentCell
        
            'Find change in price
            ClosePrice = Cells(i, 6)
            ChangeInPrice = ClosePrice - OpenPrice
            
            'Print change in price
            Cells(PrintRow, ChangeInPriceCol).Value = Round(ChangeInPrice, 2)
            
            'Find percent change in price
                If OpenPrice = 0 Then
                    PctChangeInPrice = 0
                
                Else: PctChangeInPrice = (ChangeInPrice / OpenPrice)
            
                End If
                
            
            'Print percent change in price
            Cells(PrintRow, PctChangeCol).Value = FormatPercent(PctChangeInPrice)
                
                'Color code to show change in price
                If (ChangeInPrice > 0) Then
                    'Green
                    Cells(PrintRow, ChangeInPriceCol).Interior.ColorIndex = Green
                ElseIf (ChangeInPrice <= 0) Then
                    'Red
                    Cells(PrintRow, ChangeInPriceCol).Interior.ColorIndex = Red
                End If
            'Open Price of next ticker symbol
            OpenPrice = Cells(i + 1, 3).Value
            
            'Print to next
            PrintRow = PrintRow + 1
        
        End If
        
    Next i
        Range("O2").Value = "Greatest % Increase"
        Range("O3").Value = "Greatest % Decrease"
        Range("O4").Value = "Greatest Total Volume"
        
        Range("Q2").Value = WorksheetFunction.Max(Range("K2:K" & lastrow)) * 100
        Range("Q3").Value = WorksheetFunction.Min(Range("K2:K" & lastrow)) * 100
        Range("Q4").Value = WorksheetFunction.Max(Range("L2:L" & lastrow))
     
        MaxTicker = WorksheetFunction.Match(WorksheetFunction.Max(Range("K2:K" & lastrow)), Range("K2:K" & lastrow), 0)
        Range("P2") = Cells(MaxTicker + 1, 9)
        
        Dim MinTicker As Double
        
        MinTicker = Application.WorksheetFunction.Match(WorksheetFunction.Min(Range("K2:K" & lastrow)), Range("K2:K" & lastrow), 0)
        Range("P3") = Cells(MinTicker + 1, 9)
        
        GrtVolTicker = Application.Match(WorksheetFunction.Max(Range("L2:L" & lastrow)), Range("L2:L" & lastrow), 0)
        Range("P4") = Cells(GrtVolTicker + 1, 9)
        
    End Sub
        
  

