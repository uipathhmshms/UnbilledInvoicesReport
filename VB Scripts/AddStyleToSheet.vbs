Sub AddStyleToSheet()
	' ------------------------------------------------Variables declaration------------------------------------------------
    Dim tableStart As Range
    Dim firstRowRange As Range
    Dim intTableWidth As Integer
    Dim lastRow As Long
    Dim usedRange As Range
    
	' ------------------------------------------------Settings------------------------------------------------
    ' Get the used range in the worksheet
    Set usedRange = ActiveSheet.UsedRange
    
    lastRow = usedRange.Rows.Count
	' Get the width (number of columns) of the used range
    'intTableWidth = usedRange.Columns.Count
	intTableWidth=ActiveSheet.Cells(lastRow, Columns.Count).End(xlToLeft).Column

    ' Define the starting cell of the table (A1 in this case)
    Set tableStart = Range("A1")

    ' Set the range for the first row of the table based on the table's width (number of columns)
    Set firstRowRange = tableStart.Resize(1, intTableWidth)
    
	' Get the last used row in the sheet 
	'------------------------------------------------Method calls------------------------------------------------
    ' Changes the background and text colors
    ApplyFirstRowBackgroundColor firstRowRange
    
    ' Center the text in the first row
    CenterTextInFirstRow firstRowRange
    
    ' Add a filter to the first row
    AddFilterToFirstRow firstRowRange
    
    ' Freeze the first row
    FreezeFirstRow
    
    ' Automatically adjust the width of each column to fit the content of the first row
    AutoFitColumns firstRowRange, intTableWidth
  
	' Aligns all cells to center and column E and C to the right
    AlignText
	
	' Merge so the manager name will be written once
	MergeFirstColumnRowsExceptFirstAndLast
	
	' Apply background color to the last row
    ApplyLastRowBackgroundColor lastRow, intTableWidth

	' Formats all numbers with comma separator and no decimal points	
	FormatNumbers
	
	' Applys table-like styling (borders, alternating row colors, etc.)
	AddTableStyle
End Sub

Sub MergeFirstColumnRowsExceptFirstAndLast()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim firstRow As Long
    Dim mergeRange As Range
    
    ' Set the worksheet to the active sheet (modify as needed)
    Set ws = ActiveSheet
    
    ' Define the range of rows to merge
    firstRow = 2 ' Skip the first row (header)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row ' Get the last row with data in column A
    
    ' Ensure there are at least 3 rows to work with
    If lastRow <= firstRow Then
        MsgBox "Not enough rows to merge.", vbExclamation
        Exit Sub
    End If
    
    ' Define the range to merge
    Set mergeRange = ws.Range(ws.Cells(firstRow, 1), ws.Cells(lastRow - 1, 1))
    
    ' Merge rows in the first column from the second row to the penultimate row
    mergeRange.Merge
    
    ' Align the text in the merged cell
    With mergeRange
        .HorizontalAlignment = xlCenter ' Center horizontally
        .VerticalAlignment = xlTop ' Align text to the top
    End With
End Sub

' Changes the background and text colors
Sub ApplyFirstRowBackgroundColor(firstRowRange As Range)
    ' Apply the background color to the first row of the table
    firstRowRange.Interior.Color = RGB(51, 51, 51) ' RGB color for dark grey
	' Set the text color to white for the first row
    firstRowRange.Font.Color = RGB(255, 255, 255) ' RGB color for white text
End Sub

Sub ApplyLastRowBackgroundColor(lastRow As Long, intTableWidth As Integer)
    Dim ws As Worksheet
    Dim lastRowRange As Range
    
    ' Set the worksheet to the active sheet
    Set ws = ActiveSheet
    
    ' Define the range for the last row
    Set lastRowRange = ws.Range(ws.Cells(lastRow, 1), ws.Cells(lastRow, intTableWidth))
    
    ' Apply a background color to the last row (you can change the color if needed)
    lastRowRange.Interior.Color = RGB(51, 51, 51) ' RGB color for dark grey
End Sub

Sub AddFilterToFirstRow(firstRowRange As Range)
    ' Add a filter to the first row (autofilter)
    firstRowRange.AutoFilter
End Sub

Sub FreezeFirstRow()
    ' Freeze the first row
    ActiveWindow.SplitRow = 1
    ActiveWindow.FreezePanes = True
End Sub

Sub AutoFitColumns(firstRowRange As Range, intTableWidth As Integer)
    ' Automatically adjust the width of each column to fit the content of the first row
    firstRowRange.EntireColumn.AutoFit
End Sub

Sub CenterTextInFirstRow(firstRowRange As Range)
    ' Center the text in the first row
    firstRowRange.HorizontalAlignment = xlCenter
End Sub

' Aligns all cells to center and column E and C to the right
Sub AlignText()
    ' Get the used range in the active sheet
    Dim usedRange As Range
    Set usedRange = ActiveSheet.UsedRange
    
    ' Center align the text in the entire used range
    usedRange.HorizontalAlignment = xlCenter
	' Right align text in columns C and E
    
	' Right align columns C and E (except the header row)
    Dim i As Long
    For i = 2 To usedRange.Rows.Count
        Cells(i, 3).HorizontalAlignment = xlRight  ' Column C
        Cells(i, 5).HorizontalAlignment = xlRight  ' Column E
    Next i
End Sub

' Applies table-like styling (borders, alternating row colors, etc.)
Sub AddTableStyle()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    Dim headerLastCol As Long
    Dim dataLastCol As Long
    Dim tableRange As Range
    Dim i As Long
    
    ' Set the worksheet to the active sheet
    Set ws = ActiveSheet
    
    ' Get the last row
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ' Get the last column by checking both header and data
    headerLastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    dataLastCol = ws.Cells(lastRow, ws.Columns.Count).End(xlToLeft).Column
    
    ' Use the larger of the two values
    lastCol = IIf(headerLastCol > dataLastCol, headerLastCol, dataLastCol)
    
    ' Define the table range
    Set tableRange = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))
    
    ' Apply borders to the entire table
    tableRange.Borders(xlEdgeBottom).LineStyle = xlContinuous
    tableRange.Borders(xlEdgeRight).LineStyle = xlContinuous
    tableRange.Borders(xlEdgeLeft).LineStyle = xlContinuous
    tableRange.Borders(xlEdgeTop).LineStyle = xlContinuous
    tableRange.Borders(xlInsideVertical).LineStyle = xlContinuous
    tableRange.Borders(xlInsideHorizontal).LineStyle = xlContinuous
    
    ' Apply light grey borders
    tableRange.Borders.Color = RGB(211, 211, 211) ' Light grey borders
    
    ' Apply white text color to all cells in the table
    tableRange.Font.Color = RGB(255, 255, 255) ' White text
    
    ' Apply #121212 background color to all cells in the table
    tableRange.Interior.Color = RGB(18, 18, 18) ' Dark background (#121212)
End Sub

' Formats all numbers with comma separator and no decimal points
Sub FormatNumbers()
    ' Get the used range in the active sheet
    Dim usedRange As Range
    Set usedRange = ActiveSheet.UsedRange
    
    Dim cell As Range
    
    ' Loop through each cell in the used range
    For Each cell In usedRange
        If IsNumeric(cell.Value) Then
			' If the cell contains a number, format as a number with comma separator and no decimal points
			cell.NumberFormat = "#,##0" ' Number format with comma separators
        End If
    Next cell
End Sub
