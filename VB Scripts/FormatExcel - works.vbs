Option Explicit

Sub FormatExcel()
    Dim ws As Worksheet
    Dim grandTotal As Double
    Dim objRange As Range
    Dim months As Collection
    Dim lastRow As Long

    Set ws = ThisWorkbook.Sheets(1)
    Set objRange = ws.UsedRange
    Set months = New Collection

    ExtractUniqueMonths ws, objRange, months
    AddMonthHeaders ws, months

    grandTotal = 0
    ProcessRows ws, objRange, months, grandTotal

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    AddGrandTotalRow ws, lastRow, months, grandTotal

	' Remove the "סכום כולל במטבע עסקה" column (assuming it's column 6)
    'ws.Columns(6).Delete
	' Remove the "תאריך פעילות" column (assuming it's column 7)
    'ws.Columns(7).Delete
End Sub

Sub ExtractUniqueMonths(ws As Worksheet, objRange As Range, ByRef months As Collection)
    Dim i As Long
    Dim activityDate As Date
    Dim monthYear As String
    On Error Resume Next

    For i = 2 To objRange.Rows.Count
        If ws.Cells(i, 7).Value <> "" Then
            activityDate = ParseIsraeliDate(ws.Cells(i, 7).Value)
            monthYear = Format(activityDate, "mmm-yy")
'MsgBox monthYear
            months.Add monthYear, monthYear
        End If
    Next i
    On Error GoTo 0
End Sub

Sub AddMonthHeaders(ws As Worksheet, months As Collection)
    Dim i As Integer
    Dim month As Variant
    Dim headerColumn As Integer
    headerColumn = 8

    For Each month In months
        ws.Cells(1, headerColumn).Value = month
        headerColumn = headerColumn + 1
    Next month
End Sub

Sub ProcessRows(ws As Worksheet, objRange As Range, months As Collection, ByRef grandTotal As Double)
    Dim i As Long
    Dim activityDate As Date
    Dim monthYear As String
    Dim columnIndex As Integer
    Dim amount As Double

    For i = 2 To objRange.Rows.Count
        amount = ws.Cells(i, 6).Value
        If ws.Cells(i, 7).Value <> "" Then
            activityDate = ParseIsraeliDate(ws.Cells(i, 7).Value)
            monthYear = Format(activityDate, "mmm-yy")

            columnIndex = GetMonthColumnIndex(months, monthYear)

            If columnIndex > 0 Then
                ws.Cells(i, 7 + columnIndex).Value = amount
                grandTotal = grandTotal + amount
            End If
        End If
    Next i
End Sub

Function GetMonthColumnIndex(months As Collection, monthYear As String) As Integer
    Dim i As Integer
    Dim month As Variant
    i = 1

    For Each month In months
        If month = monthYear Then
            GetMonthColumnIndex = i
            Exit Function
        End If
        i = i + 1
    Next month
    GetMonthColumnIndex = -1
End Function

Sub AddGrandTotalRow(ws As Worksheet, lastRow As Long, months As Collection, grandTotal As Double)
    Dim j As Long
    Dim lastMonthColumn As Integer

    lastMonthColumn = 6 + months.Count - 1

    ws.Cells(lastRow, 1).Value = "Grand Total"
    
    For j = 6 To lastMonthColumn
        ws.Cells(lastRow, j).Formula = "=SUM(" & ws.Cells(2, j).Address & ":" & ws.Cells(lastRow - 1, j).Address & ")"
    Next

    ws.Cells(lastRow, lastMonthColumn + 1).Value = grandTotal
End Sub

Function ParseIsraeliDate(dateString As String) As Date
    Dim dateParts() As String
    dateParts = Split(dateString, "/")
    If UBound(dateParts) = 2 Then
        ParseIsraeliDate = DateSerial(CInt(dateParts(2)), CInt(dateParts(1)), CInt(dateParts(0)))
    Else
        ParseIsraeliDate = CDate("1/1/1900") ' Or some other default date
    End If
End Function
