Sub stocks()

For Each ws In Worksheets
    ' Easy Variables
    Dim Volume As Double
    Dim trow As Double
    Dim vrow As Double
    trow = 2
    vrow = 2
    Volume = 0
    
    ' Moderate Variables
    Dim open_price As Double
    Dim close_price As Double
    
    ' Hard Variables
    Dim large_check As Double
    Dim small_check As Double
    Dim volume_check As Double
    Dim gticker As String
    Dim lticker As String
    Dim vticker As String
    large_check = 0
    small_check = 0
    volume_check = 0

    ws.Range("J1").Value = "Ticker"
    ws.Range("M1").Value = "Total Volume"

    ws.Range("K1").Value = "Yearly Change"
    ws.Range("L1").Value = "Percentage Change"

    For i = 2 To (ws.Cells(Rows.Count, 1).End(xlUp).Row + 1)

       If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
       
       'Recognizes & prints first Ticker / assigns first volume / assigns first open_price
       'increases ticker row
        If ws.Cells(i - 1, 1).Value = "<ticker>" Then
          ws.Cells(trow, 10).Value = ws.Cells(i, 1).Value
          Volume = ws.Cells(i, 7).Value
          trow = trow + 1
          open_price = ws.Cells(i, 3).Value
        'All other ticker switches - prints old volume, prints new ticker
        'increases ticker & volume row
        'Obtains close price, compares with open before resetting
        Else
          close_price = ws.Cells(i - 1, 6).Value
          ws.Cells(vrow, 13).Value = Volume
          ws.Cells(vrow, 11).Value = close_price - open_price
          ws.Cells(vrow, 12).Value = ws.Cells(vrow, 11).Value / (open_price + 0.00001)
          ws.Cells(trow, 10).Value = ws.Cells(i, 1).Value
          trow = trow + 1
          vrow = vrow + 1
          Volume = ws.Cells(i, 7).Value
          open_price = ws.Cells(i, 3).Value
       End If
       'Same Ticker / Adds to rolling volume
       Else
        Volume = Volume + ws.Cells(i, 7).Value
       End If
    Next i
       
    ' Hardest
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest & Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    
    
    For i = 2 To (ws.Cells(Rows.Count, 12).End(xlUp).Row)
        If ws.Cells(i, 12).Value > large_check Then
            large_check = ws.Cells(i, 12).Value
            gticker = ws.Cells(i, 10).Value
        ElseIf ws.Cells(i, 12).Value < small_check Then
            small_check = ws.Cells(i, 12).Value
            lticker = ws.Cells(i, 10).Value
        End If
    Next i
    
    ws.Range("P2").Value = gticker
    ws.Range("Q2").Value = large_check
    ws.Range("P3").Value = lticker
    ws.Range("Q3").Value = small_check
    
    For i = 2 To (ws.Cells(Rows.Count, 13).End(xlUp).Row)
        If ws.Cells(i, 13).Value > volume_check Then
            volume_check = ws.Cells(i, 13).Value
            vticker = ws.Cells(i, 10).Value
        End If
    Next i
    
    ws.Range("P4").Value = vticker
    ws.Range("Q4").Value = volume_check
    
        ' Formatting
    For i = 2 To (ws.Cells(Rows.Count, 11).End(xlUp).Row)
        ws.Cells(i, 12).NumberFormat = "0.0000000%"
        If ws.Cells(i, 11).Value >= 0 Then
            ws.Cells(i, 11).Interior.ColorIndex = 4
        Else
            ws.Cells(i, 11).Interior.ColorIndex = 3
        End If
    Next i
    ws.Range("Q2:Q3").NumberFormat = "0.00%"
       
Next ws
End Sub
