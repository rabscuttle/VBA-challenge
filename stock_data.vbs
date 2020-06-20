Sub stocks()

    'Declare variables
    Dim lastRow As Long
    Dim ticker As String
    Dim tableRow As Integer
    
    'Loop through Worksheets
    For Each ws In Worksheets
    
        'Get Last Row
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        'Loop through rows
        tableRow = 2
        For i = 2 To lastRow
            'Check current record vs next record
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ticker = ws.Cells(i, 1).Value
                ws.Cells(tableRow, 9).Value = ticker
                tableRow = tableRow + 1
            End If
        Next i
    Next

End Sub