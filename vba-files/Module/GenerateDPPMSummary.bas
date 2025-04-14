Attribute VB_Name = "GenerateDPPMSummary"
Option Explicit
' This module generates daily, weekly, And monthly summaries of DPPM data from the "dppm-database" sheet.

Private Sub GenerateSummaryByType(summaryType As String, sourceSheetName As String, targetSheetName As String)
    Dim wsSource As Worksheet, wsTarget As Worksheet
    Dim lastRow As Long, targetRow As Long
    Dim summaryDict As Object
    Dim key As Variant
    Dim dateValue As Date
    Dim overallQuantity As Double, overallRejects As Double, inspectedQuantity As Double, inspectedRejects As Double
    Dim overallDPPM As Double, inspectedDPPM As Double
    Dim tempArray As Variant
    Dim i As Long
    Dim outputArr() As Variant

    ' Validate input parameters
    On Error Resume Next
    Set wsSource = ThisWorkbook.Sheets(sourceSheetName)
    On Error Goto 0
        If wsSource Is Nothing Then
            MsgBox "Source sheet '" & sourceSheetName & "' Not found!", vbExclamation
         Exit Sub
        End If

        ' Create Or clear the target sheet
        On Error Resume Next
        Set wsTarget = ThisWorkbook.Sheets(targetSheetName)
        On Error Goto 0
            If wsTarget Is Nothing Then
                Set wsTarget = ThisWorkbook.Sheets.Add
                wsTarget.Name = targetSheetName
            Else
                wsTarget.Cells.Clear
            End If

            ' Initialize dictionary For aggregation
            Set summaryDict = CreateObject("Scripting.Dictionary")

            ' Find the last row in the source sheet
            lastRow = wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).Row

            ' Loop through each row in the source sheet
            For i = 2 To lastRow
                dateValue = wsSource.Cells(i, 1).Value
                overallQuantity = wsSource.Cells(i, 5).Value
                overallRejects = wsSource.Cells(i, 6).Value
                inspectedQuantity = wsSource.Cells(i, 8).Value
                inspectedRejects = wsSource.Cells(i, 9).Value

                ' Determine the key based on summary type
                Select Case summaryType
                 Case "Daily"
                    key = Format(dateValue, "yyyy-mm-dd")
                 Case "Weekly"
                    ' Use ISO week number For consistency
                    ' The "yyyy" format returns the year (4 digits)
                    key = Format(dateValue, "yyyy") & "-WW" & Format(Application.WorksheetFunction.WeekNum(dateValue, 2), "00")
                 Case "Monthly"
                    ' add year To the month For uniqueness
                    key = Format(dateValue, "yyyy-mmmm") ' "yyyy-mm" format For month
                 Case Else
                    MsgBox "Invalid summary type!", vbExclamation
                 Exit Sub
                End Select

                ' Aggregate data
                If Not summaryDict.exists(key) Then
                    summaryDict.Add key, Array(0, 0, 0, 0)
                End If
                tempArray = summaryDict(key)
                tempArray(0) = tempArray(0) + overallQuantity
                tempArray(1) = tempArray(1) + overallRejects
                tempArray(2) = tempArray(2) + inspectedQuantity
                tempArray(3) = tempArray(3) + inspectedRejects
                summaryDict(key) = tempArray
            Next i

            ' Pre-size the output array: [1 To rowCount, 1 To 10]
            ReDim outputArr(1 To summaryDict.Count + 2, 1 To 10)

            ' Write headers To the target sheet
            Debug.Print "Writing headers To the target sheet."
            wsTarget.Cells.Clear
            outputArr(1, 1) = summaryType & " Summary"
            outputArr(2, 1) = summaryType
            outputArr(2, 2) = "Overall Qty Received"
            outputArr(2, 3) = "Overall Units Reject"
            outputArr(2, 4) = "Overall DPPM"
            outputArr(2, 5) = "Inspected Qty Received"
            outputArr(2, 6) = "Inspected Units Reject"
            outputArr(2, 7) = "Inspected DPPM"

            ' Write aggregated data To the target sheet
            targetRow = 3
            For Each key In summaryDict.keys
                tempArray = summaryDict(key)
                overallDPPM = 0
                inspectedDPPM = 0
                If tempArray(0) > 0 Then overallDPPM = (tempArray(1) / tempArray(0)) * 1000000
                    If tempArray(2) > 0 Then inspectedDPPM = (tempArray(3) / tempArray(2)) * 1000000
                        outputArr(targetRow, 1) = key
                        outputArr(targetRow, 2) = tempArray(0)
                        outputArr(targetRow, 3) = tempArray(1)
                        outputArr(targetRow, 4) = Format(overallDPPM, "0")
                        outputArr(targetRow, 5) = tempArray(2)
                        outputArr(targetRow, 6) = tempArray(3)
                        outputArr(targetRow, 7) = Format(inspectedDPPM, "0")

                        targetRow = targetRow + 1
                    Next key

                    ' Write all at once To the sheet, starting at row 1
                    wsTarget.Range("A1").Resize(summaryDict.Count + 2, 10).Value = outputArr

                    ' Put borders around the cells
                    With wsTarget.Range("A1:G" & targetRow - 1).Borders
                        .LineStyle = xlContinuous
                        .Weight = xlThin
                        .ColorIndex = 0
                    End With

                    ' Autofit columns
                    wsTarget.Columns("A:G").AutoFit

                    Debug.Print summaryType & " summary generated successfully!"
End Sub

Public Sub GenerateSummary()
    ' Add status bar message
    Application.StatusBar = "Generating DPPM summaries..."
    Application.StatusBar = "Generating Daily  summary..."
    GenerateSummaryByType "Daily", "dppm-database", "DailySummary"

    Application.StatusBar = "Generating Weekly summary..."
    GenerateSummaryByType "Weekly", "dppm-database", "WeeklySummary"

    Application.StatusBar = "Generating Monthly summary..."
    GenerateSummaryByType "Monthly", "dppm-database", "MonthlySummary"
End Sub
