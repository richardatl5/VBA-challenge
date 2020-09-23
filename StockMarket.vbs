Sub StockMarket()


    For Each ws In Worksheets

        Dim Ticker As String
        Dim Stock_Volume As Double
        Dim YearlyChange As Double
        Dim PercentChange As Double

        Dim Summary_Table_Row As Long
        Summary_Table_Row = 2
        Dim StartingPrice As Double
        Dim EndingPrice As Double
        Dim LastRow As Long

        Dim rowCount As Long
        rowCount = 2

        StartingPrice = 0
        EndingPrice = 0

        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Changed"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(2, 15).Value = "Lowest"
        ws.Cells(4, 15).Value = "Highest"

        LastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row

            For i = 2 To LastRow

                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

                    Ticker = ws.Cells(i, 1).Value
                    rowCount = rowCount - (rowCount - 1)
                    StartingPrice = ws.Cells(i - rowCount, 6).Value
                    EndingPrice = ws.Cells(i, 6).Value
                    YearlyChange = EndingPrice - StartingPrice
                    Stock_Volume = Stock_Volume + ws.Cells(i, 7).Value
                    PercentChange = EndingPrice / StartingPrice

                    ws.Range("I" & Summary_Table_Row).Value = Ticker
                    ws.Range("L" & Summary_Table_Row).Value = Stock_Volume
                    ws.Range("J" & Summary_Table_Row).Value = YearlyChange
                    ws.Range("K" & Summary_Table_Row).Value = PercentChange

                    Summary_Table_Row = Summary_Table_Row + 1
                    

                    Stock_Volume = 0
                    rowCount = 0

                Else

                    Stock_Volume = Stock_Volume + ws.Cells(i, 7).Value
                    rowCount = rowCount + 1

                End If
            Next i

                For J = 2 To LastRow
                If ws.Range("K" & J).Value > ws.Range("P2").Value Then
                    ws.Range("Q2").Value = ws.Range("K" & J).Value
                    ws.Range("P2").Value = ws.Range("I" & J).Value
                End If

                If ws.Range("K" & J).Value < ws.Range("P3").Value Then
                    ws.Range("Q3").Value = ws.Range("K" & J).Value
                    ws.Range("P3").Value = ws.Range("I" & J).Value
                End If

                If ws.Range("L" & J).Value > ws.Range("P4").Value Then
                    ws.Range("Q4").Value = ws.Range("L" & J).Value
                    ws.Range("P4").Value = ws.Range("I" & J).Value
                End If

            Next J

        Next ws

End Sub
