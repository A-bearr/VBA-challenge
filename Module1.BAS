Attribute VB_Name = "Module1"
Sub AnalyzeStockData()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim ticker As String
    Dim openPrice As Double
    Dim closePrice As Double
    Dim volume As Double
    Dim quarterlyChange As Double
    Dim percentageChange As Double
    Dim outputRow As Long
    
    Dim greatestIncrease As Double
    Dim greatestDecrease As Double
    Dim greatestVolume As Double
    Dim tickerGreatestIncrease As String
    Dim tickerGreatestDecrease As String
    Dim tickerGreatestVolume As String
   
    greatestIncrease = -1 * (10 ^ 10)
    greatestDecrease = 10 ^ 10
    greatestVolume = 0

    ' Loop through each worksheet
    For Each ws In ThisWorkbook.Worksheets
        ' Find the last row with data in column A
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

        ' Prepare output headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Quarterly Change"
        ws.Cells(1, 11).Value = "Percentage Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"

        outputRow = 2

        ' Loop through each row of data
        For i = 2 To lastRow
            ticker = ws.Cells(i, 1).Value
            openPrice = ws.Cells(i, 3).Value
            closePrice = ws.Cells(i, 6).Value
            volume = ws.Cells(i, 7).Value

            ' Calculate quarterly change and percentage change
            quarterlyChange = closePrice - openPrice
            If openPrice <> 0 Then
                percentageChange = (quarterlyChange / openPrice) * 100
            Else
                percentageChange = 0
            End If

            ' Output the results
            ws.Cells(outputRow, 9).Value = ticker
            ws.Cells(outputRow, 10).Value = quarterlyChange
            ws.Cells(outputRow, 11).Value = percentageChange
            ws.Cells(outputRow, 12).Value = volume

            ' Track greatest increase, decrease, and volume
            If percentageChange > greatestIncrease Then
                greatestIncrease = percentageChange
                tickerGreatestIncrease = ticker
            End If
            
            If percentageChange < greatestDecrease Then
                greatestDecrease = percentageChange
                tickerGreatestDecrease = ticker
            End If
            
            If volume > greatestVolume Then
                greatestVolume = volume
                tickerGreatestVolume = ticker
            End If

            outputRow = outputRow + 1
        Next i

        ' Apply conditional formatting for quarterly change
        With ws.Range(ws.Cells(2, 10), ws.Cells(outputRow - 1, 10)) ' Range for Quarterly Change
            .FormatConditions.Delete
            ' Format for positive change
            .FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:=0
            .FormatConditions(.FormatConditions.Count).Interior.Color = RGB(0, 255, 0) ' Green
            ' Format for negative change
            .FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:=0
            .FormatConditions(.FormatConditions.Count).Interior.Color = RGB(255, 0, 0) ' Red
        End With

        ' Output the results for greatest increase, decrease, and volume
        ws.Cells(outputRow, 15).Value = "Greatest % Increase"
        ws.Cells(outputRow, 16).Value = tickerGreatestIncrease
        ws.Cells(outputRow, 17).Value = greatestIncrease

        outputRow = outputRow + 1
        ws.Cells(outputRow, 15).Value = "Greatest % Decrease"
        ws.Cells(outputRow, 16).Value = tickerGreatestDecrease
        ws.Cells(outputRow, 17).Value = greatestDecrease

        outputRow = outputRow + 1
        ws.Cells(outputRow, 15).Value = "Greatest Total Volume"
        ws.Cells(outputRow, 16).Value = tickerGreatestVolume
        ws.Cells(outputRow, 17).Value = greatestVolume
        

        ' Reset variables for the next worksheet
        greatestIncrease = -1 * (10 ^ 10)
        greatestDecrease = 10 ^ 10
        greatestVolume = 0
    Next ws

    MsgBox "Analysis Complete!"
End Sub
