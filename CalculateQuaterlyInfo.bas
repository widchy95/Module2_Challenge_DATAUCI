Attribute VB_Name = "Module1"
Sub CalculateQuarterlyInfo()
    Dim ws As Worksheet
    Dim sheetNames As Variant
    sheetNames = Array("Q1", "Q2", "Q3", "Q4")
    
    Dim lastRow As Long
    Dim ticker As String
    Dim openPrice As Double, closePrice As Double
    Dim totalVolume As Double
    Dim startRow As Long, endRow As Long
    Dim quarter As String
    Dim outputRow As Long
    Dim i As Long, j As Integer
    
    Dim maxPercentIncrease As Double
    Dim maxPercentIncreaseTicker As String
    Dim maxPercentDecrease As Double
    Dim maxPercentDecreaseTicker As String
    Dim maxTotalVolume As Double
    Dim maxTotalVolumeTicker As String
    
    For j = LBound(sheetNames) To UBound(sheetNames)
        Set ws = ThisWorkbook.Sheets(sheetNames(j))
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        
        ' Add headers for the output
        ws.Cells(1, 10).Value = "Ticker"
        ws.Cells(1, 11).Value = "Quarterly Change"
        ws.Cells(1, 12).Value = "Percent Change"
        ws.Cells(1, 13).Value = "Total Stock Volume"
        
        outputRow = 2
        maxPercentIncrease = -1000
        maxPercentDecrease = 1000
        maxTotalVolume = 0
        
        For i = 2 To lastRow
            ticker = ws.Cells(i, 1).Value
            quarter = Year(ws.Cells(i, 2).Value) & " Q" & Application.WorksheetFunction.RoundUp(Month(ws.Cells(i, 2).Value) / 3, 0)
            
            startRow = i
            Do While ws.Cells(i, 1).Value = ticker And Year(ws.Cells(i, 2).Value) & " Q" & Application.WorksheetFunction.RoundUp(Month(ws.Cells(i, 2).Value) / 3, 0) = quarter
                i = i + 1
            Loop
            endRow = i - 1
            
            openPrice = ws.Cells(startRow, 3).Value
            closePrice = ws.Cells(endRow, 6).Value
            totalVolume = Application.WorksheetFunction.Sum(ws.Range(ws.Cells(startRow, 7), ws.Cells(endRow, 7)))
            
            ' Output the results
            ws.Cells(outputRow, 10).Value = ticker
            ws.Cells(outputRow, 11).Value = closePrice - openPrice
            ws.Cells(outputRow, 12).Value = Format(((closePrice - openPrice) / openPrice) * 100, "0.00") & "%"
            ws.Cells(outputRow, 13).Value = totalVolume
            
            ' Apply color formatting
            If ws.Cells(outputRow, 11).Value > 0 Then
            ws.Cells(outputRow, 11).Interior.Color = vbGreen
            ElseIf ws.Cells(outputRow, 11).Value < 0 Then
            ws.Cells(outputRow, 11).Interior.Color = vbRed
            End If

            
            
            ' Determine the greatest percent increase and decrease
            Dim percentChange As Double
            percentChange = ((closePrice - openPrice) / openPrice) * 100
            
            If percentChange > maxPercentIncrease Then
                maxPercentIncrease = percentChange
                maxPercentIncreaseTicker = ticker
            End If
            
            If percentChange < maxPercentDecrease Then
                maxPercentDecrease = percentChange
                maxPercentDecreaseTicker = ticker
            End If
            
            ' Determine the greatest total volume
            If totalVolume > maxTotalVolume Then
                maxTotalVolume = totalVolume
                maxTotalVolumeTicker = ticker
            End If
            
            outputRow = outputRow + 1
            i = i - 1
        Next i
        
      
        ws.Cells(2, 15).Value = "Ticker"
        ws.Cells(2, 16).Value = "Value"
        ws.Cells(3, 14).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = maxPercentIncreaseTicker
        ws.Cells(3, 16).Value = Format(maxPercentIncrease, "0.00") & "%"
        ws.Cells(4, 14).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = maxPercentDecreaseTicker
        ws.Cells(4, 16).Value = Format(maxPercentDecrease, "0.00") & "%"
        ws.Cells(5, 14).Value = "Greatest Total Volume"
        ws.Cells(5, 15).Value = maxTotalVolumeTicker
        ws.Cells(5, 16).Value = maxTotalVolume
    Next j
End Sub



