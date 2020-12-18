Attribute VB_Name = "Module2"
Sub Bonus()

'Define all variables
Dim GITicker As String
Dim GDTicker As String
Dim GVTicker As String
Dim GreatPctIncr As Double
Dim GreatPctDecr As Double
Dim GreatVol As Double
Dim ws As Worksheet
Dim currentTicker As String
Dim currentValue As Double

    ' set the counter variables to zero for each worksheet
    For Each ws In Worksheets
    GreatPctIncr = 0
    GreatPctDecr = 0
    GreatVol = 0

    'create title labels for the summary table
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"

    'tells computer to start at the end of the worksheet and find the ninth row's ("I") column with data to begin counting
    RowCount = ws.Cells(Rows.Count, "I").End(xlUp).Row

    For i = 2 To RowCount
        currentTicker = ws.Cells(i, 9).Value
        Value = ws.Cells(i, 11).Value
        
        'comparing and keeping track of the greatest percent decrease values
        If Value < GreatPctDecr Then
            GreatPctDecr = Value
            GDTicker = currentTicker
        End If
  
        'comparing and keeping track of the greatest percent increase values
        Value = ws.Cells(i, 11).Value
        If Value > GreatPctIncr Then
            GreatPctIncr = Value
            GITicker = currentTicker
        End If

        'comparing and keeping track of the greatest total volume values
        Value = ws.Cells(i, 12).Value
        If Value > GreatVol Then
            GreatVol = Value
            GVTicker = currentTicker
        End If
    Next i

    'place results into the same worksheet next to the summary table data
    ws.Range("P2").Value = GITicker
    ws.Range("Q2").Value = GreatPctIncr
    ws.Range("P3").Value = GDTicker
    ws.Range("Q3").Value = GreatPctDecr
    ws.Range("P4").Value = GVTicker
    ws.Range("Q4").Value = GreatVol

    Next ws

End Sub

