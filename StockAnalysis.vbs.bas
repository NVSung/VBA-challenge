Attribute VB_Name = "Module1"
Sub StockAnalysis()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim ticker As String
    Dim openPrice As Double
    Dim closePrice As Double
    Dim totalVolume As Double
    Dim quarterlyChange As Double
    Dim percentChange As Double
    Dim startRow As Long
    Dim endRow As Long
    Dim outputRow As Long
    Dim greatestIncrease As Double
    Dim greatestDecrease As Double
    Dim greatestVolume As Double
    Dim greatestIncreaseTicker As String
    Dim greatestDecreaseTicker As String
    Dim greatestVolumeTicker As String

    ' Loop in each worksheet
    For Each ws In ThisWorkbook.Worksheets
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        outputRow = 2 ' Start from the second row as the first row is for headers

        ' Add headers
        ws.Cells(1, 10).Value = "Ticker"
        ws.Cells(1, 11).Value = "Quarterly Change"
        ws.Cells(1, 12).Value = "Percentage Change"
        ws.Cells(1, 13).Value = "Total Stock Volume"

        ' Implement tracking variables
        greatestIncrease = -1
        greatestDecrease = 1
        greatestVolume = 0
        startRow = 2

        ' Loop through data rows
        Do While startRow <= lastRow
            ticker = ws.Cells(startRow, 1).Value
            openPrice = ws.Cells(startRow, 3).Value
            totalVolume = 0
            endRow = startRow

            ' Sum volume and find the end row for the current ticker
            Do While endRow <= lastRow And ws.Cells(endRow, 1).Value = ticker
                totalVolume = totalVolume + ws.Cells(endRow, 7).Value
                endRow = endRow + 1
            Loop

            endRow = endRow - 1
            closePrice = ws.Cells(endRow, 6).Value

            ' Calculate changes
            quarterlyChange = closePrice - openPrice
            If openPrice <> 0 Then
                percentChange = (quarterlyChange / openPrice) * 100
            Else
                percentChange = 0
            End If

            ' Output results
            ws.Cells(outputRow, 10).Value = ticker
            ws.Cells(outputRow, 11).Value = quarterlyChange
            ws.Cells(outputRow, 12).Value = Format(percentChange, "0.00") & "%" ' Changing the format to 2 decimal places
            ws.Cells(outputRow, 13).Value = totalVolume

            ' Apply conditional formatting
            If quarterlyChange > 0 Then
                ws.Cells(outputRow, 11).Interior.Color = RGB(0, 255, 0)
            ElseIf quarterlyChange < 0 Then
                ws.Cells(outputRow, 11).Interior.Color = RGB(255, 0, 0)
            End If

            ' Track greatest changes and volume
            If percentChange > greatestIncrease Then
                greatestIncrease = percentChange
                greatestIncreaseTicker = ticker
            End If

            If percentChange < greatestDecrease Then
                greatestDecrease = percentChange
                greatestDecreaseTicker = ticker
            End If

            If totalVolume > greatestVolume Then
                greatestVolume = totalVolume
                greatestVolumeTicker = ticker
            End If

            outputRow = outputRow + 1
            startRow = endRow + 1
        Loop

        ' Output greatest values
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(2, 16).Value = greatestIncreaseTicker
        ws.Cells(2, 17).Value = Format(greatestIncrease, "0.00") & "%"

        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(3, 16).Value = greatestDecreaseTicker
        ws.Cells(3, 17).Value = Format(greatestDecrease, "0.00") & "%"

        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(4, 16).Value = greatestVolumeTicker
        ws.Cells(4, 17).Value = greatestVolume
    Next ws
End Sub
