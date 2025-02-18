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
	' Remove the "סכום כולל במטבע עסקה" column (assuming it's column 6)
    ws.Columns(6).Delete
	' Remove the "תאריך פעילות" column (assuming it's column 7)
	ws.Columns(6).Delete
	' Merge duplicate rows based on "שם לקוח" and "קןד סעיף תקציבי"
    MergeDuplicateRows ws
	' Add grand total
	lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    AddGrandTotalRow ws, lastRow, months, grandTotal
	' Set excel data direction from right to left
	SetSheetDirectionRTL ws
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

    ' Determine the column where the last month's data is
    lastMonthColumn = 6 + months.Count - 1

    ' Set "Grand Total" in the last row, first column
    ws.Cells(lastRow, 1).Value = "Grand Total"
    
    ' Loop through the columns and calculate the SUM for each month
    For j = 6 To lastMonthColumn
        ws.Cells(lastRow, j).Formula = "=SUM(" & ws.Cells(2, j).Address & ":" & ws.Cells(lastRow - 1, j).Address & ")"
    Next j

    ' Set the grand total value in the last column (after the last month)
    ws.Cells(lastRow, lastMonthColumn + 1).Value = grandTotal

    ' Add "Grand Total" header in the first row, last month column + 1
    ws.Cells(1, lastMonthColumn + 1).Value = "Grand Total"
    
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

Sub SetSheetDirectionRTL(sheet As Worksheet)
    With sheet
        .DisplayRightToLeft = True ' Set sheet direction to Right-to-Left
        .Cells.HorizontalAlignment = xlRight ' Align text to the right
    End With
End Sub

Sub MergeDuplicateRows(ws As Worksheet)
    Dim lastRow As Long
    Dim i As Long, j As Long
    Dim dict As Object
    Dim key As Variant ' Key must be a Variant for For Each loop
    Dim rowValues As Variant
    Dim columnCount As Integer
    Dim newRow As Variant

    Set dict = CreateObject("Scripting.Dictionary")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    columnCount = ws.UsedRange.Columns.Count

    ' Loop through rows and consolidate data
    For i = 2 To lastRow
        key = ws.Cells(i, 3).Value & "_" & ws.Cells(i, 4).Value ' Ensure קוד סעיף תקציבי is part of the key

        If dict.exists(key) Then
            rowValues = dict(key)
            ' Sum numeric values for the months while keeping empty cells blank
            For j = 6 To columnCount
                If IsNumeric(ws.Cells(i, j).Value) And ws.Cells(i, j).Value <> "" Then
                    If rowValues(1, j) = "" Then
                        rowValues(1, j) = ws.Cells(i, j).Value ' Keep original value if empty
                    Else
                        rowValues(1, j) = rowValues(1, j) + ws.Cells(i, j).Value ' Sum values if both are numbers
                    End If
                End If
            Next j
            dict(key) = rowValues
        Else
            ' Store the entire row in a dictionary (as a 2D array)
            rowValues = ws.Rows(i).Value
            dict.Add key, rowValues
        End If
    Next i

    ' Clear old data
    ws.Rows("2:" & lastRow).ClearContents

    ' Write back merged data
    i = 2
    For Each key In dict.keys ' Ensure key is a Variant
        newRow = dict(key)
        ws.Rows(i).Value = newRow
        i = i + 1
    Next key
End Sub

