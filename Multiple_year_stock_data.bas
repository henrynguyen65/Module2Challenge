Attribute VB_Name = "Module1"
Sub MultipleYearStockAnalysis()
    
    Dim Ticker As String
    Dim OpeningPrice As Double
    Dim ClosingPrice As Double
    Dim YearlyChange As Double
    Dim PercentChange As Double
    Dim TotalVolume As Double
    
    Dim GreatestPercentIncrease As Double
    Dim GreatestPercentDecrease As Double
    Dim GreatestTotalVolume As Double
    
    Dim GreatestPercentIncreaseTicker As String
    Dim GreatestPercentDecreaseTicker As String
    Dim GreatestTotalVolumeTicker As String
    
    Dim SumRow As Long
    Dim ws As Worksheet
    Dim LastRow As Long
    
    For Each ws In ThisWorkbook.Worksheets
        
        YearlyChange = 0
        TotalVolume = 0
        GreatestIncrease = 0
        GreatestDecrease = 0
        GreatestVolume = 0
        SumRow = 2
        StartingPriceRow = 2
        
        LastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        
        
    For i = 2 To LastRow
            
        If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                
            Ticker = ws.Cells(i, 1).Value
            OpeningPrice = ws.Cells(StartingPriceRow, 3).Value
            StartingPriceRow = i + 1
            ClosingPrice = ws.Cells(i, 6).Value
                
            YearlyChange = ClosingPrice - OpeningPrice
            
        If OpeningPrice <> 0 Then
        
            PercentChange = YearlyChange / OpeningPrice
            
        Else
            
            PercentChange = 0
            
        End If

            ws.Cells(SumRow, 9).Value = Ticker
            ws.Cells(SumRow, 10).Value = YearlyChange
            ws.Cells(SumRow, 11).Value = PercentChange
            ws.Cells(SumRow, 12).Value = TotalVolume
                
            If YearlyChange > 0 Then
                ws.Cells(SumRow, 10).Interior.Color = RGB(0, 255, 0)
            ElseIf YearlyChange < 0 Then
                ws.Cells(SumRow, 10).Interior.Color = RGB(255, 0, 0)
            
            End If
                
            If PercentChange > GreatestPercentIncrease Then
                GreatestPercentIncrease = PercentChange
                GreatestPercentIncreaseTicker = Ticker
            End If
            
            If PercentChange < GreatestPercentDecrease Then
                GreatestPercentDecrease = PercentChange
                GreatestPercentDecreaseTicker = Ticker
            End If
            
            If TotalVolume > GreatestTotalVolume Then
                GreatestTotalVolume = TotalVolume
                GreatestTotalVolumeTicker = Ticker
            End If

                YearlyChange = 0
                TotalVolume = 0
                SumRow = SumRow + 1
        Else
                
            TotalVolume = TotalVolume + ws.Cells(i, 7).Value
                                             
    End If
    Next i
        ws.Range("K2:K" & SumRow).NumberFormat = "0.00%"
        ws.Range("Q2:Q3").NumberFormat = "0.00%"
        
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
       
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 16).Value = GreatestPercentIncreaseTicker
        ws.Cells(3, 16).Value = GreatestPercentDecreaseTicker
        ws.Cells(4, 16).Value = GreatestTotalVolumeTicker
        ws.Cells(2, 17).Value = GreatestPercentIncrease
        ws.Cells(3, 17).Value = GreatestPercentDecrease
        ws.Cells(4, 17).Value = GreatestTotalVolume
    Next ws
End Sub


