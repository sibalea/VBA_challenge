Sub CalculateQuarterlyVolumeAndChange_Revue()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim ticker As String
    Dim openPrice As Double
    Dim closePrice As Double
    Dim change As Double
    Dim percentChange As Double
    Dim totalVolume As Double ' Changed to Double to handle large numbers
    Dim summaryRow As Long
    Dim firstRowIndex As Long
    Dim lastRowIndex As Long
    Dim i As Long
   
    ' Initialize summary table headers
    Set wb = ThisWorkbook
   
    ' Loop through each worksheet
    For Each ws In wb.Worksheets
        ' Clear existing summary if it exists
        ws.Range("I:L").ClearContents
       
        ' Write summary table headers
        ws.Cells(1, 9).Value = "Tickers"
        ws.Cells(1, 10).Value = "Quarterly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Volume"
       
        ' Determine the last row of data in column A
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
       
        ' Determine the last row index based on the quarter
        Select Case ws.Name
            Case "Q1"
                lastRowIndex = 62
            Case "Q2"
                lastRowIndex = 63
            Case "Q3", "Q4"
                lastRowIndex = 64
            Case Else
                ' Default to the last row in column A if sheet name doesn't match
                lastRowIndex = lastRow
        End Select
       
        ' Initialize first row index
        firstRowIndex = 2
       
        ' Initialize summary table starting row
        summaryRow = 2
       
        ' Reset totalVolume for each sheet
        totalVolume = 0
       
        ' Loop through each row to calculate total volume and change in price
        For i = 2 To lastRow
            ' Check if the ticker changes
            If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
                ' Calculate change in price if it's the last row of the quarter or the ticker changes
                If i > firstRowIndex Then
                    openPrice = ws.Cells(firstRowIndex, 3).Value ' Open price on the first day
                    closePrice = ws.Cells(i - 1, 6).Value ' Close price on the last day of the previous ticker
                    change = closePrice - openPrice
                    If openPrice <> 0 Then
                        percentChange = (change / openPrice)
                    Else
                        percentChange = 0 ' Handle division by zero if openPrice is zero
                    End If
                   
                    ' Write summary values to summary table
                    ws.Cells(summaryRow, 9).Value = ticker ' Column J
                    ws.Cells(summaryRow, 10).Value = change ' Column K
                    ws.Cells(summaryRow, 11).Value = percentChange ' Column L
                    ws.Cells(summaryRow, 12).Value = totalVolume ' Column M (total volume)
                    summaryRow = summaryRow + 1
                End If
               
                ' Reset variables for the next ticker
                ticker = ws.Cells(i, 1).Value
                totalVolume = 0
                firstRowIndex = i
            End If
           
            ' Accumulate total volume for the current ticker
            totalVolume = totalVolume + ws.Cells(i, 7).Value ' Accumulate total volume
           
            ' Handle the last ticker in the sheet
            If i = lastRow Then
                openPrice = ws.Cells(firstRowIndex, 3).Value ' Open price on the first day
                closePrice = ws.Cells(i, 6).Value ' Close price on the last day of the last ticker
                change = closePrice - openPrice
                If openPrice <> 0 Then
                    percentChange = (change / closePrice)
                Else
                    percentChange = 0 ' Handle division by zero if openPrice is zero
                End If
               
                ' Write summary values to summary table for the last ticker
                ws.Cells(summaryRow, 9).Value = ticker ' Column I
                ws.Cells(summaryRow, 10).Value = change ' Column J
                ws.Cells(summaryRow, 11).Value = percentChange ' Column K
                ws.Cells(summaryRow, 12).Value = totalVolume ' Column L (total volume)
                summaryRow = summaryRow + 1
            End If
            
        Next i
    Next ws
End Sub

Sub greatest_value_revue()
Dim wb As Workbook
    Dim quaterlyMax As Double
    Dim quaterlyMin As Double
    Dim totalVolumeMax As Double
    Dim lastRow As Long
    Dim tickerMax As String
    Dim tickerMin As String
    Dim tickerVolumeMax As String
    
    ' Initialize summary table headers
    Set wb = ThisWorkbook
    
    For Each ws In wb.Worksheets
        ' Clear existing summary if it exists
        ws.Range("P:R").ClearContents
       
        ' Write summary table headers
        ws.Cells(1, 17).Value = "Ticker"
        ws.Cells(1, 18).Value = "Value"
       
        ' Determine the last row of data in column J
        lastRow = ws.Cells(ws.Rows.Count, "J").End(xlUp).Row
        
        ' Initialize values
        quaterlyMax = ws.Cells(2, 11).Value
        quaterlyMin = ws.Cells(2, 11).Value
        totalVolumeMax = ws.Cells(2, 12).Value
        
        ' Loop through each row
        For i = 3 To lastRow
            If ws.Cells(i, 11).Value > quaterlyMax Then
                quaterlyMax = ws.Cells(i, 11).Value
                tickerMax = ws.Cells(i, 9).Value
            ElseIf ws.Cells(i, 11).Value < quaterlyMin Then
                quaterlyMin = ws.Cells(i, 11).Value
                tickerMin = ws.Cells(i, 9).Value
            ElseIf ws.Cells(i, 12).Value > totalVolumeMax Then
                totalVolumeMax = ws.Cells(i, 12).Value
                tickerVolumeMax = ws.Cells(i, 9).Value
            End If
        Next i
        
        ' Write summary values to summary table for the last ticker
        ws.Cells(2, 16).Value = "Greatest % increase" ' Cell P2
        ws.Cells(3, 16).Value = "Greatest % decrease" ' Cell P3
        ws.Cells(4, 16).Value = "Greatest Total Volume" ' Cell P4
        
        ws.Cells(2, 17).Value = tickerMax ' Cell Q2
        ws.Cells(3, 17).Value = tickerMin ' Cell Q3
        ws.Cells(4, 17).Value = tickerVolumeMax ' Cell Q4
        
        ws.Cells(2, 18).Value = quaterlyMax ' Cell R2
        ws.Cells(3, 18).Value = quaterlyMin ' Cell R3
        ws.Cells(4, 18).Value = totalVolumeMax ' Cell R4
        
    Next ws
                    
End Sub

Sub color_coded()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim i As Long
        
    Set wb = ThisWorkbook
    For Each ws In wb.Worksheets
        'show column 10 value in green if positive, and red if negative
        'i could be set to 1500 but I prefer the entire column
        'just in case some other values get added to it
        For i = 2 To 96000
            j = 10
            If ws.Cells(i, "J").Value > 0 Then
                ws.Cells(i, "J").Interior.Color = RGB(0, 255, 0) ' Green
            Else: ws.Cells(i, "J").Interior.Color = RGB(255, 0, 0) ' Red
            End If
        Next i
    Next ws
End Sub


