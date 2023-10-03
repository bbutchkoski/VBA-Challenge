Attribute VB_Name = "Module1"
Sub Multiple_Year_Stock_Data()

        
    Dim ws As Worksheet
    Dim Ticker As String
    Dim YearlyChange As Double
    Dim PercentChange As Double
    Dim TotalVolume As Double
    Dim LastRow As Long
    Dim SummaryRow As Long
    Dim OpeningPrice As Double
    Dim ClosingPrice As Double
    
  For Each ws In Worksheets
    ws.Cells(1, 9) = "Ticker"
    ws.Cells(1, 10) = "Yearly Change"
    ws.Cells(1, 11) = "Percent Change"
    ws.Cells(1, 12) = "Total Stock Volume"
    ws.Cells(1, 16) = "Ticker"
    ws.Cells(1, 17) = "Value"
    ws.Cells(2, 15) = "Greatest % Value"
    ws.Cells(3, 15) = "Greatest % value"
    ws.Cells(4, 15) = "Greatest Total Value"
    
    'LastRow = ws.Cells(Rows.Count, 1).End(x1Up).Row'
    LastRow = Cells(Rows.Count, "A").End(xlUp).Row
    
  
    SummaryRow = 2
    
    
    For i = 2 To LastRow
    
        
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                Ticker = ws.Cells(i, 1).Value
                
                OpeningPrice = ws.Cells(i, 3).Value
                
                ClosingPrice = ws.Cells(i, 6).Value
               
                YearlyChange = ClosingPrice - OpeningPrice
                
                If OpeningPrice <> 0 Then
                    PercentChange = (YearlyChange / OpeningPrice) * 100
                Else
                    PercentChange = 0
                End If
                
                ws.Cells(SummaryRow, 9).Value = Ticker
                ws.Cells(SummaryRow, 10).Value = YearlyChange
                ws.Cells(SummaryRow, 11).Value = PercentChange
                ws.Cells(SummaryRow, 12).Value = TotalVolume
                'ws.Cells(SummaryRow, 12).Value = Application.WorksheetFunction.Sum(ws.Range(ws.Cells(i - TotalVolume + 1, 7), ws.Cells(i, 7)'
                
                ws.Cells(SummaryRow, 11).NumberFormat = "0.00%"
                
                If YearlyChange > 0 Then
                    ws.Cells(SummaryRow, 10).Interior.Color = RGB(0, 255, 0)
                ElseIf YearlyChange < 0 Then
                    ws.Cells(SummaryRow, 10).Interior.Color = RGB(255, 0, 0)
                End If
                
                SummaryRow = SummaryRow + 1
                
                TotalVolume = 0
            End If
            
            TotalVolume = TotalVolume + ws.Cells(i, 7).Value
        Next i
    
        Next ws
        'SummaryLastRow = ws.Cells(Rows.Count, 9).EndIxlUp.Row'
        
        'Dim MaxPercentIncrease As Double'
        'Dim MaxPercent'
        
                          

End Sub

