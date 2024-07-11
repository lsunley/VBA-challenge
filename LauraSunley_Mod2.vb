Sub CalculateStockMetricsAllSheets()
    Dim ws As Worksheet
    Dim lastRow, uniqueRow, i As Long
    Dim uniqueTickers As Collection
    Dim ticker As Variant
    Dim firstOpen, lastClose As Double
    Dim tickerVolume As Double
    Dim minDate, maxDate As Date
    Dim data, results As Variant
    
    Application.ScreenUpdating = False
    
    For Each ws In ThisWorkbook.Worksheets
        ' Find the last row with data in column A
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        
        ' Read data into an array
        data = ws.Range("A2:G" & lastRow).Value
        
        ' Create a collection of unique tickers
        Set uniqueTickers = New Collection
        On Error Resume Next
        For i = 1 To UBound(data)
            uniqueTickers.Add data(i, 1), CStr(data(i, 1))
        Next i
        On Error GoTo 0
        
        ' Prepare results array
        ReDim results(1 To uniqueTickers.Count, 1 To 4)
        
        ' Add headers in I1:L1
        With ws
            .Range("I1").Value = "Ticker"
            .Range("J1").Value = "Change"
            .Range("K1").Value = "%"
            .Range("L1").Value = "Volume"
        End With
        
        ' Process each unique ticker
        uniqueRow = 1
        For Each ticker In uniqueTickers
            ' Initialize variables
            firstOpen = 0
            lastClose = 0
            minDate = DateValue("9999-12-31")
            maxDate = DateValue("1900-01-01")
            tickerVolume = 0
            
            ' Loop through the data array to find the required values
            For i = 1 To UBound(data)
                If data(i, 1) = ticker Then
                    ' Update the min date and first open price
                    If data(i, 2) < minDate Then
                        minDate = data(i, 2)
                        firstOpen = data(i, 3)
                    End If
                    
                    ' Update the max date and last close price
                    If data(i, 2) > maxDate Then
                        maxDate = data(i, 2)
                        lastClose = data(i, 6)
                    End If
                    
                    ' Sum the total volume
                    tickerVolume = tickerVolume + data(i, 7)
                End If
            Next i
            
            ' Store results
            results(uniqueRow, 1) = ticker
            results(uniqueRow, 2) = lastClose - firstOpen
            If firstOpen <> 0 Then
                ' Calculate percentage change
                results(uniqueRow, 3) = (lastClose - firstOpen) / Abs(firstOpen) ' Calculate percentage change as a decimal
            Else
                results(uniqueRow, 3) = 0
            End If
            results(uniqueRow, 4) = tickerVolume
            
            ' Apply conditional formatting to column J (Change)
            If results(uniqueRow, 2) > 0 Then
                ws.Cells(uniqueRow + 1, 10).Interior.Color = RGB(146, 208, 80) ' Green color for positive change
            ElseIf results(uniqueRow, 2) < 0 Then
                ws.Cells(uniqueRow + 1, 10).Interior.Color = RGB(255, 0, 0) ' Red color for negative change
            End If
            
            ' Apply conditional formatting to column K (%)
            If results(uniqueRow, 3) > 0 Then
                ws.Cells(uniqueRow + 1, 11).Interior.Color = RGB(146, 208, 80) ' Green color for positive %
            ElseIf results(uniqueRow, 3) < 0 Then
                ws.Cells(uniqueRow + 1, 11).Interior.Color = RGB(255, 0, 0) ' Red color for negative %
            End If
            
            uniqueRow = uniqueRow + 1
        Next ticker
        
        ' Write results to the worksheet
        ws.Range("I2:L" & UBound(results, 1) + 1).Value = results
        
        ' Format column K as percentage
        ws.Range("K2:K" & UBound(results, 1) + 1).NumberFormat = "0.00%"
        
        ' Adjusting the actual values in column K to reflect percentage correctly
        For i = 2 To UBound(results, 1) + 1
            ws.Cells(i, 11).Value = results(i - 1, 3)
        Next i
        
        ' Identify the stock with the greatest % increase, greatest % decrease, and greatest total volume
        Dim greatestIncrease, greatestDecrease, greatestVolume As Double
        Dim increaseTicker, decreaseTicker, volumeTicker As String
        
        greatestIncrease = -1
        greatestDecrease = 1
        greatestVolume = 0
        
        For i = 2 To UBound(results, 1) + 1
            If ws.Cells(i, 11).Value > greatestIncrease Then
                greatestIncrease = ws.Cells(i, 11).Value
                increaseTicker = ws.Cells(i, 9).Value
            End If
            If ws.Cells(i, 11).Value < greatestDecrease Then
                greatestDecrease = ws.Cells(i, 11).Value
                decreaseTicker = ws.Cells(i, 9).Value
            End If
            If ws.Cells(i, 12).Value > greatestVolume Then
                greatestVolume = ws.Cells(i, 12).Value
                volumeTicker = ws.Cells(i, 9).Value
            End If
        Next i
        
        ' Output the results
        ws.Range("N1").Value = "Metrics"
        ws.Range("N2").Value = "Greatest % Increase"
        ws.Range("O2").Value = increaseTicker
        ws.Range("P2").Value = greatestIncrease
        ws.Range("P2").NumberFormat = "0.00%"
        
        ws.Range("N3").Value = "Greatest % Decrease"
        ws.Range("O3").Value = decreaseTicker
        ws.Range("P3").Value = greatestDecrease
        ws.Range("P3").NumberFormat = "0.00%"
        
        ws.Range("N4").Value = "Greatest Total Volume"
        ws.Range("O4").Value = volumeTicker
        ws.Range("P4").Value = greatestVolume
        
    Next ws

    Application.ScreenUpdating = True
End Sub

