Sub FormatExcel()
    ' Variables declaration
    Dim ws As Worksheet
    Dim grandTotal As Double
    Dim objRange As Range
    Dim earliestDate As Date
    Dim activityDate As Date
    Dim startMonth As Integer
    Dim startYear As Integer
    Dim months(11) As String ' Array to store month names
    Dim i As Integer, j As Integer
    Dim columnIndex As Integer
    Dim projectManager As String
    Dim chargeMethod As String
    Dim clientName As String
    Dim budgetCode As String
    Dim budgetItem As String
    Dim amount As Double
    Dim lastRow As Long

    ' Set the active worksheet
    Set ws = ThisWorkbook.Sheets(1)

    ' Get the range of the data (Assuming data starts at row 2)
    Set objRange = ws.UsedRange

    ' Find the earliest date from the "תאריך פעילות" column (column 7)
    earliestDate = ws.Cells(2, 7).Value
    For i = 3 To objRange.Rows.Count
        activityDate = ws.Cells(i, 7).Value
        If activityDate < earliestDate Then
            earliestDate = activityDate
        End If
    Next

    ' Get the starting month and year based on the earliest date
    startMonth = Month(earliestDate)
    startYear = Year(earliestDate)

    ' Generate month names from the starting month
    For i = 0 To 11
        months(i) = Format(DateSerial(startYear, startMonth + i, 1), "mmm-yy")
        If startMonth + i > 12 Then
            startYear = startYear + 1
            startMonth = startMonth + i - 12
        End If
    Next

    ' Add column headers for months dynamically based on the earliest date
    For i = 0 To UBound(months)
        ws.Cells(1, 6 + i).Value = months(i)
    Next

    ' Initialize Grand Total
    grandTotal = 0

    ' Process each row of data
    For i = 2 To objRange.Rows.Count
        ' Read the necessary values from each row
        projectManager = ws.Cells(i, 1).Value
        chargeMethod = ws.Cells(i, 2).Value
        clientName = ws.Cells(i, 3).Value
        budgetCode = ws.Cells(i, 4).Value
        budgetItem = ws.Cells(i, 5).Value
        amount = ws.Cells(i, 6).Value
        activityDate = ws.Cells(i, 7).Value
        
        ' Get the month and year from the activity date
        activityMonth = Month(activityDate)
        activityYear = Year(activityDate)
        
        ' Calculate the column index based on the dynamic month
        columnIndex = -1
        For j = 0 To UBound(months)
            If Month(DateSerial(activityYear, activityMonth, 1)) = Month(DateSerial(startYear, startMonth + j, 1)) Then
                columnIndex = j
                Exit For
            End If
        Next
        
        ' Add the amount to the respective month's column
        ws.Cells(i, columnIndex + 6).Value = amount

        ' Add the value to the Grand Total
        grandTotal = grandTotal + amount
    Next

    ' Add the Grand Total row in the first available empty row
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    ws.Cells(lastRow, 5).Value = "Grand Total"

    ' Sum up all months for the Grand Total Row
    For j = 6 To 6 + UBound(months) ' Columns based on the dynamic month range
        ws.Cells(lastRow, j).Formula = "=SUM(" & ws.Cells(2, j).Address & ":" & ws.Cells(lastRow - 1, j).Address & ")"
    Next

    ' Optional: Add the grand total in the last cell for easy reference
    ws.Cells(lastRow, 6 + UBound(months) + 1).Value = grandTotal

End Sub
