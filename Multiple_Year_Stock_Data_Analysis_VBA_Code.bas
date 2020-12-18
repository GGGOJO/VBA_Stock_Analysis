Attribute VB_Name = "Module1"
Sub StockMarketAnalysis()

'Define all variables
Dim i As Long
Dim j As Integer
Dim total As Double
Dim change As Double
Dim start As Long
Dim percentChange As Double
Dim ws As Worksheet

    'set the counter valriables to zero for each worksheet
    For Each ws In Worksheets
    j = 0
    change = 0
    total = 0
    start = 2

    'create title labels for the summary table
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"

    'tells computer to start at the end of the worksheet and find the first row with data in the first column "A" (ticker) to begin counting
    RowCount = ws.Cells(Rows.Count, "A").End(xlUp).Row

    'create a conditional for loop that goes from first row of data (that's row 2) through the last row of data (that's RowCount)
    For i = 2 To RowCount

    'create conditional statement to compare a row with the one above it since we start from the bottom
    'the next if statement addresses the rows with bad data (don't include data with 0 volume cells)
    'start finding the non-zero total stock volume for each ticker and store results in variable
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            total = total + ws.Cells(i, 7).Value

            If total = 0 Then
                ws.Range("I" & 2 + j).Value = Cells(i, 1).Value
                ws.Range("J" & 2 + j).Value = 0
                ws.Range("K" & 2 + j).Value = "%" & 0
                ws.Range("L" & 2 + j).Value = 0
            Else
                If ws.Cells(start, 3) = 0 Then
                    For findValue = start To i
                        If ws.Cells(findValue, 3).Value <> 0 Then
                            start = findValue
                            Exit For
                        End If
                    Next findValue
                End If

            'next conditional to find the yearly change and the percent change
            change = (ws.Cells(i, 6) - ws.Cells(start, 3))
            percentChange = Round((change / ws.Cells(start, 3) * 100), 2)

            'begin on the next stock ticker
            start = i + 1

            'place results into the same worksheet next to the raw data
            ws.Range("I" & 2 + j).Value = ws.Cells(i, 1).Value
            ws.Range("J" & 2 + j).Value = Round(change, 2)
            ws.Range("K" & 2 + j).Value = "%" & percentChange
            ws.Range("L" & 2 + j).Value = total

                'color code the yearly change column that ended the year positive (green is 4), negative (red is 3), and no change (white is 0)
                If change > 0 Then
                    ws.Range("J" & 2 + j).Interior.ColorIndex = 4
                ElseIf change < 0 Then
                    ws.Range("J" & 2 + j).Interior.ColorIndex = 3
                Else
                    ws.Range("J" & 2 + j).Interior.ColorIndex = 0
                End If

            End If

            'reset the variables to zero to count for a new stock ticker
            j = j + 1
            total = 0
            change = 0

        Else
            total = total + ws.Cells(i, 7).Value
        End If

    Next i

Next ws

End Sub


