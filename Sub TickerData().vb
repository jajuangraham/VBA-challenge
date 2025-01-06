Sub TickerData()
    Dim ws As Worksheet
    Dim Ticker As String
    Dim QuarterStartPrice As Double
    Dim QuarterEndPrice As Double
    Dim TotalVolume As Double
    Dim RowCount As Long
    Dim SummaryRow As Integer
    Dim GreatestIncrease As Double
    Dim GreatestDecrease As Double
    Dim GreatestVolume As Double
    Dim GreatestIncreaseTicker As String
    Dim GreatestDecreaseTicker As String
    Dim GreatestVolumeTicker As String

    For Each ws In ThisWorkbook.Worksheets
        ws.Activate
        RowCount = Cells(Rows.Count, 1).End(xlUp).Row
        SummaryRow = 2
        TotalVolume = 0
        GreatestIncrease = 0
        GreatestDecrease = 0
        GreatestVolume = 0

        ' Add headers to the summary table
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Quarterly Change"
        Cells(1, 11).Value = "% Change"
        Cells(1, 12).Value = "Total Volume"

        For i = 2 To RowCount
            TotalVolume = TotalVolume + Cells(i, 7).Value

            ' Check for ticker change
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Or i = RowCount Then
                Ticker = Cells(i, 1).Value
                QuarterStartPrice = Cells(i - (i - 2), 3).Value
                QuarterEndPrice = Cells(i, 6).Value

                ' Calculate metrics
                QuarterlyChange = QuarterEndPrice - QuarterStartPrice
                If QuarterStartPrice <> 0 Then
                    PercentChange = (QuarterlyChange / QuarterStartPrice) * 100
                Else
                    PercentChange = 0
                End If

                ' Populate the summary table
                Cells(SummaryRow, 9).Value = Ticker
                Cells(SummaryRow, 10).Value = QuarterlyChange
                Cells(SummaryRow, 11).Value = PercentChange
                Cells(SummaryRow, 12).Value = TotalVolume

                ' Apply conditional formatting
                If QuarterlyChange > 0 Then
                    Cells(SummaryRow, 10).Interior.Color = RGB(0, 255, 0) ' Green for positive
                Else
                    Cells(SummaryRow, 10).Interior.Color = RGB(255, 0, 0) ' Red for negative
                End If

                ' Track greatest values
                If PercentChange > GreatestIncrease Then
                    GreatestIncrease = PercentChange
                    GreatestIncreaseTicker = Ticker
                End If

                If PercentChange < GreatestDecrease Then
                    GreatestDecrease = PercentChange
                    GreatestDecreaseTicker = Ticker
                End If

                If TotalVolume > GreatestVolume Then
                    GreatestVolume = TotalVolume
                    GreatestVolumeTicker = Ticker
                End If

                ' Reset for next ticker
                TotalVolume = 0
                SummaryRow = SummaryRow + 1
            End If
        Next i

        ' Output the greatest values
        Cells(2, 14).Value = "Greatest % Increase"
        Cells(2, 15).Value = GreatestIncreaseTicker
        Cells(2, 16).Value = GreatestIncrease

        Cells(3, 14).Value = "Greatest % Decrease"
        Cells(3, 15).Value = GreatestDecreaseTicker
        Cells(3, 16).Value = GreatestDecrease

        Cells(4, 14).Value = "Greatest Volume"
        Cells(4, 15).Value = GreatestVolumeTicker
        Cells(4, 16).Value = GreatestVolume
    Next ws

    
End Sub
