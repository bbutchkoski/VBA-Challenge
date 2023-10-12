Attribute VB_Name = "Module1"
Sub Multiple_Year_Stock_Data()




    For Each ws In Worksheets



    Dim Ticker As String
    Dim YearlyChange As Double
    Dim PercentChange As Double
    Dim TotalVolume As Double
    Dim LastRow As Long
    Dim SummaryRow As Long
    Dim OpeningPrice As Double
    Dim ClosingPrice As Double



    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"



    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row



    SummaryRow = 2



    For i = 2 To LastRow



If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

    Ticker = ws.Cells(i, 1).Value



    OpeningPrice = ws.Cells(i, 3).Value
    ClosingPrice = ws.Cells(i, 6).Value
    YearlyChange = ClosingPrice - OpeningPrice
If OpeningPrice <> 0 Then
    PercentChange = (YearlyChange / OpeningPrice)
    Else
        PercentChange = 0

            End If





    ws.Cells(SummaryRow, 9).Value = Ticker
    ws.Cells(SummaryRow, 10).Value = YearlyChange
    ws.Cells(SummaryRow, 11).Value = PercentChange
    ws.Cells(SummaryRow, 12).Value = TotalVolume

  

    ws.Cells(SummaryRow, 11).NumberFormat = "0.00%"

    If YearlyChange > 0 Then
        ws.Cells(SummaryRow, 10).Interior.ColorIndex = 4
ElseIf YearlyChange < 0 Then
        ws.Cells(SummaryRow, 10).Interior.ColorIndex = 3

    End If


    SummaryRow = SummaryRow + 1



    TotalVolume = 0
End If

    TotalVolume = TotalVolume + ws.Cells(i, 7).Value
Next i

    SummaryLastRow = ws.Cells(Rows.Count, 9).End(xlUp).Row




        Dim Greatest_Increase As Double
        Dim Greatest_Decrease As Double
        Dim Greatest_Volume As Double

        Greatest_Increase = ws.Cells(2, 11).Value
        Greatest_Decrease = ws.Cells(2, 11).Value
        Greatest_Volume = ws.Cells(2, 12).Value
        LastRow_Summary = ws.Cells(Rows.Count, 10).End(xlUp).Row

    For j = 2 To LastRow_Summary


        If ws.Cells(j, 11) > Greatest_Increase Then
        Greatest_Increase = ws.Cells(j, 11)
        ws.Cells(2, 17) = Greatest_Increase
        ws.Cells(2, 17).NumberFormat = "0.00%"
        ws.Cells(2, 16) = ws.Cells(j, 9)

    End If

    If ws.Cells(j, 11) < Greatest_Decrease Then
        Greatest_Decrease = ws.Cells(j, 11)
        ws.Cells(3, 17) = Greatest_Decrease
        ws.Cells(3, 17).NumberFormat = "0.00%"
        ws.Cells(3, 16) = ws.Cells(j, 9)

    End If

    If ws.Cells(j, 12) > Greatest_Volume Then
        Greatest_Volume = ws.Cells(j, 12)
        ws.Cells(4, 17) = Greatest_Volume
        ws.Cells(4, 16) = ws.Cells(j, 9)

    End If

        Next j


    Next ws


End Sub
  

        
    

