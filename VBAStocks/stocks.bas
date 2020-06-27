Attribute VB_Name = "Module1"
Sub stocks()

    'Declare variables
    Dim lastRow As Long
    Dim ticker As String
    Dim tableRow As Integer
    Dim priceOpen As Double
    Dim priceClose As Double
    Dim volOpen As Double
    Dim volClose As Double
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim totalVolume As Double
    Dim rg As Range
    Dim cond1 As FormatCondition
    Dim cond2 As FormatCondition
    Dim highTicker As String
    Dim lowTicker As String
    Dim totalTicker As String
    Dim highValue As Double
    Dim lowValue As Double
    Dim totalValue As Double
   
    'Loop through Worksheets
    For Each ws In Worksheets
   
        'Get Last Row
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
       
        'Set values for first stock
        priceOpen = ws.Cells(2, 3).Value
        volOpen = ws.Cells(2, 7).Value
   
        'Set column headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
       
        'Format columns
        'Set column K as percentage
        ws.Range("K2:K" & lastRow).NumberFormat = "0.00%"
        ws.Range("Q2:Q3").NumberFormat = "0.00%"
        'Conditional formatting for column J
        Set rg = ws.Range("J2:J" & lastRow)
        Set cond1 = rg.FormatConditions.Add(xlCellValue, xlGreater, "=0")
        Set cond2 = rg.FormatConditions.Add(xlCellValue, xlLess, "=0")
        With cond1
        .Interior.Color = vbGreen
        End With
        With cond2
        .Interior.Color = vbRed
        End With

        'Set variables outside of loop
        tableRow = 2
        totalVolume = 0
        highTicker = ""
        highValue = 0
        lowTicker = ""
        lowValue = 0
        totalTicker = ""
        totalValue = 0

        'Loop through rows
        For i = 2 To lastRow
            'Zero handling
            If volOpen = 0 And ws.Cells(i, 1) = ws.Cells(i + 1, 1) Then
                volOpen = ws.Cells(i + 1, 7).Value
            End If
            If priceOpen = 0 And ws.Cells(i, 1) = ws.Cells(i + 1, 1) Then
                priceOpen = ws.Cells(i + 1, 3).Value
            End If

            'Add totalVolume
            totalVolume = totalVolume + ws.Cells(i, 7).Value
            If totalVolume > totalValue Then
                totalValue = totalVolume
                totalTicker = ws.Cells(i, 1).Value
            End If
           
            'Check current record vs next record
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ticker = ws.Cells(i, 1).Value
                priceClose = ws.Cells(i, 6).Value
                volClose = ws.Cells(i, 7).Value
                'Calculate values
                yearlyChange = priceClose - priceOpen
                If priceOpen <> 0 Then
                    percentChange = yearlyChange / priceOpen
                End If
                'Set values for Final Summary table
                If percentChange > highValue Then
                    highValue = percentChange
                    highTicker = ws.Cells(i, 1).Value
                End If
                If percentChange < lowValue Then
                    lowValue = percentChange
                    lowTicker = ws.Cells(i, 1).Value
                End If
                           
                'Set table value
                ws.Cells(tableRow, 9).Value = ticker
                ws.Cells(tableRow, 10).Value = yearlyChange
                ws.Cells(tableRow, 11).Value = percentChange
                ws.Cells(tableRow, 12).Value = totalVolume
                'Advance tableRow
                tableRow = tableRow + 1
                'Set open values for next stock
                priceOpen = ws.Cells(i + 1, 3).Value
                volOpen = ws.Cells(i + 1, 7).Value
                totalVolume = 0
            End If
        Next i

        'Fill last table
        ws.Cells(2, 16).Value = highTicker
        ws.Cells(2, 17).Value = highValue
        ws.Cells(3, 16).Value = lowTicker
        ws.Cells(3, 17).Value = lowValue
        ws.Cells(4, 16).Value = totalTicker
        ws.Cells(4, 17).Value = totalValue
    Next ws

End Sub



