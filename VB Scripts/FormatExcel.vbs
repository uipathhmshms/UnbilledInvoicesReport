Sub FormatExcel()
    ' Variables declaration
    Dim ws As Worksheet
    Dim grandTotal As Double
    Dim objRange As Range
    Dim months(11) As String ' Array to store month names
    Dim startMonth As Integer
    Dim startYear As Integer
    Dim lastRow As Long

    ' Set the active worksheet
    Set ws = ThisWorkbook.Sheets(1)

    ' Get the range of the data (Assuming data starts at row 2)
    Set objRange = ws.UsedRange

    ' Find the earliest date from the "תאריך פעילות" column (column 7)
    Dim earliestDate As Date
    earliestDate = FindEarliestDate(ws, objRange)

    ' Get the starting month and year based on the earliest date
    startMonth = Month(earliestDate)
    startYear = Year(earliestDate)

    ' Generate month names from the starting month
    GenerateMonthNames startMonth, startYear, months

    ' Add column headers for months dynamically based on the earliest date
    AddMonthHeaders ws, months

    ' Initialize Grand Total
    grandTotal = 0

    ' Process each row of data
    ProcessRows ws, objRange, months, grandTotal

    ' Add the Grand Total row in the first available empty row
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    AddGrandTotalRow ws, lastRow, months, grandTotal
End Sub

' Method to find the earliest date in the "תאריך פעילות" column
Function FindEarliestDate(ws As Worksheet, objRange As Range) As Date
    Dim i As Integer
    Dim activityDate As Date
    Dim earliestDate As Date
    earliestDate = ws.Cells(2, 7).Value
    For i = 3 To objRange.Rows.Count
        activityDate = ws.Cells(i, 7).Value
        If activityDate < earliestDate Then
            earliestDate = activityDate
        End If
    Next
    FindEarliestDate = earliestDate
End Function

' Method to generate month names dynamically starting from the earliest month
Sub GenerateMonthNames(startMonth As Integer, startYear As Integer, ByRef months() As String)
    Dim i As Integer
    For i = 0 To 11
        months(i) = Format(DateSerial(startYear, startMonth + i, 1), "mmm-yy")
        If startMonth + i > 12 Then
            startYear = startYear + 1
            startMonth = startMonth + i - 12
        End If
    Next
End Sub

' Method to add the month headers in the first row of the worksheet
Sub AddMonthHeaders(ws As Worksheet, months() As String)
    Dim i As Integer
    For i = 0 To UBound(months)
        ws.Cells(1, 6 + i).Value = months(i)
    Next
End Sub

' Method to process each row of data and add amounts to the respective month columns
Sub ProcessRows(ws As Worksheet, objRange As Range, months() As String, ByRef grandTotal As Double)
    Dim i As Integer
    Dim j As Integer
    Dim projectManager As String
    Dim chargeMethod As String
    Dim clientName As String
    Dim budgetCode As String
    Dim budgetItem As String
    Dim amount As Double
    Dim activityDate As Date
    Dim columnIndex As Integer

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
        Dim activityMonth As Integer
        Dim activityYear As Integer
        activityMonth = Month(activityDate)
        activityYear = Year(activityDate)

        ' Calculate the column index based on the dynamic month
        columnIndex = activityMonth - startMonth + 1

        ' Ensure the column index wraps around properly when crossing December
        If columnIndex < 1 Then
            columnIndex = columnIndex + 12 ' Handle wraparound from January to December
        End If

        ' Add the amount to the respective month's column
        ws.Cells(i, 6 + columnIndex).Value = amount

        ' Add the value to the Grand Total
        grandTotal = grandTotal + amount
    Next
End Sub

' Method to add the Grand Total row and the grand total sum formulas
Sub AddGrandTotalRow(ws As Worksheet, lastRow As Long, months() As String, grandTotal As Double)
    Dim j As Integer
    Dim lastMonthColumn As Integer

    ' Calculate the last month column based on the dynamic months
    lastMonthColumn = 6 + UBound(months) ' The last column for the last month

    ' Add "Grand Total" in the column before the last month column
    ws.Cells(lastRow, 1).Value = "Grand Total"
    
    ' Sum up all months for the Grand Total Row
    For j = 8 To lastMonthColumn ' Columns based on the dynamic month range
        ws.Cells(lastRow, j).Formula = "=SUM(" & ws.Cells(2, j).Address & ":" & ws.Cells(lastRow - 1, j).Address & ")"
    Next

    ' Optional: Add the grand total in the last cell for easy reference
    ws.Cells(lastRow, lastMonthColumn + 1).Value = grandTotal
End Sub
