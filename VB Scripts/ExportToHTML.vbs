Sub ExportToHTML()
    Dim ws As Worksheet
    Dim htmlFile As String
    Dim fileNameWithoutExt As String
    Dim objChart As Object

    ' Set the active worksheet
    Set ws = ThisWorkbook.Sheets(1)

    ' Get the name of the workbook without the extension
    fileNameWithoutExt = Left(ThisWorkbook.Name, InStrRev(ThisWorkbook.Name, ".") - 1)

    ' Define the output HTML file path using the workbook name
    htmlFile = ThisWorkbook.Path & "\" & fileNameWithoutExt & ".html"

    ' Find and hide the chart (assuming the chart is the first chart object)
    On Error Resume Next
    Set objChart = ws.ChartObjects(1) ' Adjust the index if there are multiple charts
    If Not objChart Is Nothing Then
        objChart.Visible = False
    End If
    On Error GoTo 0

    ' Export the sheet to HTML
    ws.SaveAs Filename:=htmlFile, FileFormat:=xlHTML

    ' Make the chart visible again (if it was hidden)
    If Not objChart Is Nothing Then
        objChart.Visible = True
    End If
End Sub
