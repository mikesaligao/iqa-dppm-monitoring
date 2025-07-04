Attribute VB_Name = "GenerateDPPMSummary"
'=====================================================================================
' Module: GenerateDPPMSummary
' Purpose: Generate daily, weekly, and monthly DPPM summaries from the DPPM output table.
' Author: [Your Name]
' Date: [YYYY-MM-DD]
' Description:
'   - Aggregates DPPM data by period (daily, weekly, monthly)
'   - Writes summary tables to target sheets
'   - Handles configuration, logging, and error management
'=====================================================================================
Option Explicit

' --- Module-level Variables For Config ---
Private m_sDailySummarySheetName As String, m_sDailySummaryTableName As String
Private m_sWeeklySummarySheetName As String, m_sWeeklySummaryTableName As String
Private m_sMonthlySummarySheetName As String, m_sMonthlySummaryTableName As String
Private m_Config As Object ' Module-level config dictionary

'=====================================================================================
' Public Sub: GenerateSummary
' Purpose: Entry point to generate all DPPM summaries (daily, weekly, monthly)
'=====================================================================================
Public Sub GenerateSummary(Optional control As IRibbonControl = Nothing)
    Dim procName As String: procName = Config.PROC_GENERATE_SUMMARY
    On Error GoTo ErrorHandler
    Utils.InitStatusBar procName
    Utils.LogMessage "[" & procName & "] Starting summary generation..."
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    ' Load global configuration
    Set m_Config = Utils.GetGlobalConfig()
    If m_Config Is Nothing Or m_Config.Count = 0 Then
        Utils.LogMessage "[" & procName & "] Global configuration not loaded. Aborting.", True
        GoTo CleanUp
    End If

    ' Generate summaries
    GenerateSummaryByType "Daily"
    GenerateSummaryByType "Weekly"
    GenerateSummaryByType "Monthly"
    GenerateSupplierMonthlySummary

    Utils.LogMessage "[" & procName & "] All summaries generated."
    MsgBox "DPPM Summaries generated successfully!", vbInformation

CleanUp:
    If Utils.g_blnStatusBarActive Then Utils.ResetStatusBar procName
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Exit Sub
ErrorHandler:
    Utils.LogMessage "[" & procName & "] ERROR " & Err.Number & ": " & Err.Description, True
    If Utils.g_blnStatusBarActive Then Utils.ResetStatusBar procName, True, Err.Description
    Resume CleanUp
End Sub

'=====================================================================================
' Private Sub: GenerateSummaryByType
' Purpose: Generate summary for a specific period type (Daily, Weekly, Monthly)
'=====================================================================================
Private Sub GenerateSummaryByType(summaryType As String)
    Dim wsSource As Worksheet, wsTarget As Worksheet
    Dim tblSource As ListObject, tblTarget As ListObject
    Dim summaryDict As Object, outputArr() As Variant
    Dim targetSheetName As String, targetTableName As String
    Dim procName As String: procName = Config.PROC_GENERATE_SUMMARY_BY_TYPE & " (" & summaryType & ")"
    On Error GoTo ErrorHandler
    Utils.UpdateStatusBarMessage "Generating " & summaryType & " Summary...", True

    ' Get source sheet and table (dppm-database)
    Set wsSource = Utils.GetSheet(Config.DPPM_OUTPUT_SHEET_NAME)
    If wsSource Is Nothing Then
        Utils.LogMessage "[" & procName & "] Source sheet '" & Config.DPPM_OUTPUT_SHEET_NAME & "' not found!", True
        Exit Sub
    End If
    Set tblSource = wsSource.ListObjects(Config.DPPM_OUTPUT_TABLE_NAME)
    If tblSource Is Nothing Then
        Utils.LogMessage "[" & procName & "] Source table '" & Config.DPPM_OUTPUT_TABLE_NAME & "' not found on sheet '" & Config.DPPM_OUTPUT_SHEET_NAME & "'!", True
        Exit Sub
    End If

    ' Determine target sheet and table names based on summary type
    Select Case summaryType
    Case "Daily"
        targetSheetName = Config.CONFIG_KEY_DPPM_DAILY_SHEET_NAME
        targetTableName = Config.CONFIG_KEY_DPPM_DAILY_TABLE_NAME
    Case "Weekly"
        targetSheetName = Config.CONFIG_KEY_DPPM_WEEKLY_SHEET_NAME
        targetTableName = Config.CONFIG_KEY_DPPM_WEEKLY_TABLE_NAME
    Case "Monthly"
        targetSheetName = Config.CONFIG_KEY_DPPM_MONTHLY_SHEET_NAME
        targetTableName = Config.CONFIG_KEY_DPPM_MONTHLY_TABLE_NAME
    Case Else
        Utils.LogMessage "[" & procName & "] Invalid summary type: '" & summaryType & "'", True
        Exit Sub
    End Select

    ' Check for Excel sheet name length limit
    If Len(targetSheetName) > 31 Then
        Utils.LogMessage "[" & procName & "] Target sheet name '" & targetSheetName & "' exceeds Excel's 31 character limit!", True
        Exit Sub
    End If

    ' Get or create the target sheet
    Set wsTarget = GetOrCreateSheet(ThisWorkbook, targetSheetName)
    If wsTarget Is Nothing Then
        Utils.LogMessage "[" & procName & "] Could not get or create target sheet '" & targetSheetName & "'.", True
        Exit Sub
    End If

    ' Aggregate data for the summary type
    Set summaryDict = AggregateSummaryData(tblSource, summaryType, procName)

    ' Write summary to sheet
    outputArr = BuildSummaryOutputArray(summaryDict)
    WriteSummaryToSheet wsTarget, outputArr, targetTableName, summaryType, procName

    ' Format the summary table
    FormatSummaryTable wsTarget, targetTableName, summaryType, procName

    Utils.UpdateStatusBarMessage summaryType & " Summary generated.", stageComplete:=True
    Utils.LogMessage "[" & procName & "] " & summaryType & " summary generated successfully on sheet '" & targetSheetName & "', table '" & targetTableName & "'."
    Exit Sub
ErrorHandler:
    Utils.LogMessage "[" & procName & "] ERROR " & Err.Number & ": " & Err.Description, True
End Sub

'=====================================================================================
' Private Function: AggregateSummaryData
' Purpose: Aggregate DPPM data by summary type (Daily, Weekly, Monthly)
' Returns: Dictionary with summary data
'=====================================================================================
Private Function AggregateSummaryData(tblSource As ListObject, summaryType As String, procName As String) As Object
    Dim summaryDict As Object
    Dim sourceDataArr As Variant
    Dim lastRow As Long, i As Long
    Dim sourceColDateIdx As Long, sourceColOverallQtyIdx As Long, sourceColOverallRejectIdx As Long, sourceColInspectedQtyIdx As Long, sourceColInspectedRejectIdx As Long
    Dim dateValue As Variant, key As Variant, weekNum As Variant, yearPart As String, formattedDate As String
    Dim overallQuantity As Double, overallRejects As Double, inspectedQuantity As Double, inspectedRejects As Double
    Dim tempArray As Variant
    Set summaryDict = CreateObject("Scripting.Dictionary")

    If tblSource.ListRows.Count = 0 Then
        Set AggregateSummaryData = summaryDict
        Exit Function
    End If
    sourceDataArr = tblSource.DataBodyRange.value
    lastRow = UBound(sourceDataArr, 1)
    sourceColDateIdx = Utils.GetColumnIndexByName(tblSource, Config.DPPM_COL_DATE)
    sourceColOverallQtyIdx = Utils.GetColumnIndexByName(tblSource, Config.DPPM_COL_OVERALL_QTY)
    sourceColOverallRejectIdx = Utils.GetColumnIndexByName(tblSource, Config.DPPM_COL_OVERALL_REJECT)
    sourceColInspectedQtyIdx = Utils.GetColumnIndexByName(tblSource, Config.DPPM_COL_INSPECTED_QTY)
    sourceColInspectedRejectIdx = Utils.GetColumnIndexByName(tblSource, Config.DPPM_COL_INSPECTED_REJECT)
    If sourceColDateIdx = 0 Or sourceColOverallQtyIdx = 0 Or sourceColOverallRejectIdx = 0 Or sourceColInspectedQtyIdx = 0 Or sourceColInspectedRejectIdx = 0 Then
        Set AggregateSummaryData = summaryDict
        Exit Function
    End If
    Dim startTime As Double: startTime = Timer
    For i = 1 To lastRow
        If i Mod Utils.STATUS_BAR_RECORD_UPDATE_INTERVAL = 0 Or Timer - Utils.g_lngLastStatusBarUpdateTime > Utils.STATUS_BAR_UPDATE_INTERVAL_SECONDS Then
            Utils.UpdateStatusBarProgress "Aggregating " & summaryType, i, lastRow, startTime
        End If
        dateValue = sourceDataArr(i, sourceColDateIdx)
        If IsEmpty(dateValue) Or IsNull(dateValue) Or Not IsDate(dateValue) Then GoTo NextRow
        formattedDate = Format(dateValue, "yyyy-mm-dd")
        overallQuantity = SafeToDouble(sourceDataArr(i, sourceColOverallQtyIdx))
        overallRejects = SafeToDouble(sourceDataArr(i, sourceColOverallRejectIdx))
        inspectedQuantity = SafeToDouble(sourceDataArr(i, sourceColInspectedQtyIdx))
        inspectedRejects = SafeToDouble(sourceDataArr(i, sourceColInspectedRejectIdx))
        Select Case summaryType
        Case "Daily": key = Format(dateValue, "yyyy-mm-dd")
        Case "Weekly"
            On Error GoTo NextRow
            weekNum = DatePart("ww", dateValue, vbMonday, vbFirstFourDays)
            yearPart = Format(dateValue, "yyyy")
            key = yearPart & "-WW" & Format(weekNum, "00")
        Case "Monthly": key = Format(dateValue, "yyyy-mmmm")
        Case Else: GoTo NextRow
        End Select
        If Not summaryDict.exists(key) Then summaryDict.Add key, Array(0, 0, 0, 0)
        tempArray = summaryDict(key)
        tempArray(0) = tempArray(0) + overallQuantity
        tempArray(1) = tempArray(1) + overallRejects
        tempArray(2) = tempArray(2) + inspectedQuantity
        tempArray(3) = tempArray(3) + inspectedRejects
        summaryDict(key) = tempArray
NextRow:
    Next i
    Set AggregateSummaryData = summaryDict
End Function

'=====================================================================================
' Private Function: BuildSummaryOutputArray
' Purpose: Build output array for summary table from dictionary
'=====================================================================================
Private Function BuildSummaryOutputArray(summaryDict As Object) As Variant
    Dim outputArr() As Variant, tempArray As Variant, key As Variant
    Dim overallDPPM As Double, inspectedDPPM As Double
    Dim targetRow As Long
    ReDim outputArr(1 To summaryDict.Count + 1, 1 To 7)
    outputArr(1, 1) = Config.SUMMARY_COL_PERIOD
    outputArr(1, 2) = Config.SUMMARY_COL_OVERALL_QTY
    outputArr(1, 3) = Config.SUMMARY_COL_OVERALL_REJECT
    outputArr(1, 4) = Config.SUMMARY_COL_OVERALL_DPPM_CALC
    outputArr(1, 5) = Config.SUMMARY_COL_INSPECTED_QTY
    outputArr(1, 6) = Config.SUMMARY_COL_INSPECTED_REJECT
    outputArr(1, 7) = Config.SUMMARY_COL_INSPECTED_DPPM_CALC
    targetRow = 2
    For Each key In summaryDict.keys
        tempArray = summaryDict(key)
        overallDPPM = 0: inspectedDPPM = 0
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
    BuildSummaryOutputArray = outputArr
End Function

'=====================================================================================
' Private Sub: WriteSummaryToSheet
' Purpose: Write summary output array to worksheet and create table
'=====================================================================================
Private Sub WriteSummaryToSheet(wsTarget As Worksheet, outputArr As Variant, targetTableName As String, summaryType As String, procName As String)
    Dim dataRange As Range, tblTarget As ListObject
    wsTarget.Cells.Clear
    wsTarget.Cells(1, 1).value = summaryType & " DPPM Summary"
    wsTarget.Cells(1, 1).Font.Bold = True
    If UBound(outputArr, 1) > 1 Then
        Set dataRange = wsTarget.Range("A2").Resize(UBound(outputArr, 1), 7)
        dataRange.value = outputArr
    Else
        Set dataRange = wsTarget.Range("A2").Resize(1, 7)
        dataRange.value = Application.WorksheetFunction.Index(outputArr, 1, 0)
    End If
    On Error Resume Next
    wsTarget.ListObjects(targetTableName).Delete
    On Error GoTo 0
    Set tblTarget = wsTarget.ListObjects.Add(xlSrcRange, dataRange, , xlYes)
    tblTarget.Name = targetTableName
    tblTarget.TableStyle = Config.DEFAULT_TABLE_STYLE
    Utils.LogMessage "[" & procName & "] Created table '" & targetTableName & "' on sheet '" & wsTarget.Name & "'."
End Sub

'=====================================================================================
' Private Sub: FormatSummaryTable
' Purpose: Sort and format the summary table
'=====================================================================================
Private Sub FormatSummaryTable(wsTarget As Worksheet, targetTableName As String, summaryType As String, procName As String)


    Dim tblTarget As ListObject
    Set tblTarget = wsTarget.ListObjects(targetTableName)
    If tblTarget.ListRows.Count > 0 Then
        With tblTarget.Sort
            .SortFields.Clear
            .SortFields.Add key:=tblTarget.ListColumns(Config.SUMMARY_COL_PERIOD).Range, SortOn:=xlSortOnValues, Order:=xlAscending
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
    End If
    tblTarget.Range.Columns.AutoFit
    With tblTarget.Range
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    Utils.LogMessage "[" & procName & "] Sorted and formatted table '" & targetTableName & "'."
End Sub



'=====================================================================================
' Private Function: SafeToDouble
' Purpose: Converts a value to Double, returning 0 if not numeric or error.
'=====================================================================================
Private Function SafeToDouble(val As Variant) As Double
    On Error Resume Next
    SafeToDouble = 0
    If IsNumeric(val) Then SafeToDouble = CDbl(val)
    On Error GoTo 0
End Function

'=====================================================================================
' Private Function: LoadAndValidateSummaryConfig
' Purpose: Loads configuration specific to summary sheets/tables.
' Returns: Boolean (True if valid, False if error)
'=====================================================================================
Private Function LoadAndValidateSummaryConfig() As Boolean
    Dim procName As String: procName = Config.PROC_LOAD_SUMMARY_CONFIG
    Utils.UpdateStatusBarMessage "Loading Summary Configuration...", True
    On Error GoTo ErrorHandler
    If m_Config Is Nothing Or m_Config.Count = 0 Then
        Utils.LogMessage "[" & procName & "] Global configuration (m_Config) is not available.", True
        LoadAndValidateSummaryConfig = False
        Exit Function
    End If
    Dim requiredKeys As Variant, keyName As Variant
    requiredKeys = Array( _
                   Config.CONFIG_KEY_DPPM_DAILY_SHEET_NAME, Config.CONFIG_KEY_DPPM_DAILY_TABLE_NAME, _
                   Config.CONFIG_KEY_DPPM_WEEKLY_SHEET_NAME, Config.CONFIG_KEY_DPPM_WEEKLY_TABLE_NAME, _
                   Config.CONFIG_KEY_DPPM_MONTHLY_SHEET_NAME, Config.CONFIG_KEY_DPPM_MONTHLY_TABLE_NAME _
                                                             )
    For Each keyName In requiredKeys
        If Not m_Config.exists(keyName) Or Trim(CStr(m_Config(keyName))) = "" Then
            Utils.LogMessage "[" & procName & "] Missing or empty configuration for: '" & keyName & "' in 'Config' sheet.", True
            LoadAndValidateSummaryConfig = False
            Exit Function
        End If
    Next keyName
    m_sDailySummarySheetName = Trim(CStr(m_Config(Config.CONFIG_KEY_DPPM_DAILY_SHEET_NAME)))
    m_sDailySummaryTableName = Trim(CStr(m_Config(Config.CONFIG_KEY_DPPM_DAILY_TABLE_NAME)))
    m_sWeeklySummarySheetName = Trim(CStr(m_Config(Config.CONFIG_KEY_DPPM_WEEKLY_SHEET_NAME)))
    m_sWeeklySummaryTableName = Trim(CStr(m_Config(Config.CONFIG_KEY_DPPM_WEEKLY_TABLE_NAME)))
    m_sMonthlySummarySheetName = Trim(CStr(m_Config(Config.CONFIG_KEY_DPPM_MONTHLY_SHEET_NAME)))
    m_sMonthlySummaryTableName = Trim(CStr(m_Config(Config.CONFIG_KEY_DPPM_MONTHLY_TABLE_NAME)))
    LoadAndValidateSummaryConfig = True
    Utils.LogMessage "[" & procName & "] Summary Configuration loaded and validated successfully."
    Utils.UpdateStatusBarMessage "Summary Configuration loaded.", stageComplete:=True
    Exit Function
ErrorHandler:
    Utils.LogMessage "[" & procName & "] ERROR " & Err.Number & ": " & Err.Description, True
    LoadAndValidateSummaryConfig = False
End Function
'=====================================================================================
' The "Supplier-Monthly DPPM Summary" table is formatted as an Excel Table.
' You can use this table as a source for Power Pivot, Pivot Tables, or Slicers.
' - To use with Power Pivot: Add the table to the Data Model.
' - To use with Slicers: Insert a Pivot Table from this table, then add Slicers for "Supplier" and "Month".
'=====================================================================================
' Private Sub: GenerateSupplierMonthlySummary
' Purpose: Generate summary per Supplier per Month
'=====================================================================================
Private Sub GenerateSupplierMonthlySummary()
    Dim procName As String: procName = "GenerateSupplierMonthlySummary"
    On Error GoTo ErrorHandler
    Utils.UpdateStatusBarMessage "Generating Supplier-Monthly Summary...", True

    ' Load config
    Set m_Config = Utils.GetGlobalConfig()
    Dim sheetName As String, tableName As String
    sheetName = Config.CONFIG_KEY_DPPM_SUPPLIER_SHEET_NAME
    tableName = Config.CONFIG_KEY_DPPM_SUPPLIER_TABLE_NAME

    ' Get source table
    Dim wsSource As Worksheet, tblSource As ListObject
    Set wsSource = Utils.GetSheet(Config.DPPM_OUTPUT_SHEET_NAME)
    If wsSource Is Nothing Then
        Utils.LogMessage "[" & procName & "] Source sheet not found!", True
        Exit Sub
    End If
    Set tblSource = wsSource.ListObjects(Config.DPPM_OUTPUT_TABLE_NAME)
    If tblSource Is Nothing Then
        Utils.LogMessage "[" & procName & "] Source table not found!", True
        Exit Sub
    End If

    ' Get or create target sheet
    Dim wsTarget As Worksheet
    Set wsTarget = GetOrCreateSheet(ThisWorkbook, sheetName)
    If wsTarget Is Nothing Then
        Utils.LogMessage "[" & procName & "] Could not get or create target sheet.", True
        Exit Sub
    End If

    ' Aggregate data
    Dim summaryDict As Object
    Set summaryDict = AggregateSupplierMonthlyData(tblSource, procName)

    ' Build output array
    Dim outputArr As Variant
    outputArr = BuildSupplierMonthlyOutputArray(summaryDict)

    ' Write to sheet
    WriteSupplierMonthlySummary wsTarget, outputArr, tableName, procName

    ' Format table
    FormatSummaryTable wsTarget, tableName, "Supplier-Monthly", procName

    Utils.UpdateStatusBarMessage "Supplier-Monthly Summary generated.", stageComplete:=True
    Utils.LogMessage "[" & procName & "] Supplier-Monthly summary generated successfully."
    Exit Sub
ErrorHandler:
    Utils.LogMessage "[" & procName & "] ERROR " & Err.Number & ": " & Err.Description, True
End Sub

'=====================================================================================
' Private Function: AggregateSupplierMonthlyData
' Purpose: Aggregate data by Supplier and Month
'=====================================================================================
Private Function AggregateSupplierMonthlyData(tblSource As ListObject, procName As String) As Object
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim arr As Variant: arr = tblSource.DataBodyRange.value
    Dim lastRow As Long: lastRow = UBound(arr, 1)
    Dim idxDate As Long, idxSupplier As Long, idxOverallQty As Long, idxOverallReject As Long
    Dim idxInspectedQty As Long, idxInspectedReject As Long
    idxDate = Utils.GetColumnIndexByName(tblSource, Config.DPPM_COL_DATE)
    idxSupplier = Utils.GetColumnIndexByName(tblSource, Config.DPPM_COL_SUPPLIER)
    idxOverallQty = Utils.GetColumnIndexByName(tblSource, Config.DPPM_COL_OVERALL_QTY)
    idxOverallReject = Utils.GetColumnIndexByName(tblSource, Config.DPPM_COL_OVERALL_REJECT)
    idxInspectedQty = Utils.GetColumnIndexByName(tblSource, Config.DPPM_COL_INSPECTED_QTY)
    idxInspectedReject = Utils.GetColumnIndexByName(tblSource, Config.DPPM_COL_INSPECTED_REJECT)
    Dim i As Long, key As String, dateVal As Variant, monthKey As String, supplier As String
    Dim tempArr As Variant
    For i = 1 To lastRow
        dateVal = arr(i, idxDate)
        supplier = arr(i, idxSupplier)
        If IsDate(dateVal) And supplier <> "" Then
            monthKey = Format(dateVal, "yyyy-mmm")
            key = supplier & "|" & monthKey
            If Not dict.exists(key) Then dict.Add key, Array(0, 0, 0, 0)
            tempArr = dict(key)
            tempArr(0) = tempArr(0) + SafeToDouble(arr(i, idxOverallQty))
            tempArr(1) = tempArr(1) + SafeToDouble(arr(i, idxOverallReject))
            tempArr(2) = tempArr(2) + SafeToDouble(arr(i, idxInspectedQty))
            tempArr(3) = tempArr(3) + SafeToDouble(arr(i, idxInspectedReject))
            dict(key) = tempArr
        End If
    Next i
    Set AggregateSupplierMonthlyData = dict
End Function

'=====================================================================================
' Private Function: BuildSupplierMonthlyOutputArray
' Purpose: Build output array for Supplier-Monthly summary
'=====================================================================================
Private Function BuildSupplierMonthlyOutputArray(dict As Object) As Variant
    Dim outputArr() As Variant, tempArr As Variant, key As Variant
    Dim overallDPPM As Double, inspectedDPPM As Double
    Dim targetRow As Long, supplier As String, monthKey As String, arrSplit As Variant
    ReDim outputArr(1 To dict.Count + 1, 1 To 8)
    outputArr(1, 1) = "Supplier"
    outputArr(1, 2) = "Date"
    outputArr(1, 3) = Config.SUMMARY_COL_OVERALL_QTY
    outputArr(1, 4) = Config.SUMMARY_COL_OVERALL_REJECT
    outputArr(1, 5) = Config.SUMMARY_COL_OVERALL_DPPM_CALC
    outputArr(1, 6) = Config.SUMMARY_COL_INSPECTED_QTY
    outputArr(1, 7) = Config.SUMMARY_COL_INSPECTED_REJECT
    outputArr(1, 8) = Config.SUMMARY_COL_INSPECTED_DPPM_CALC
    targetRow = 2
    For Each key In dict.keys
        arrSplit = Split(key, "|")
        supplier = arrSplit(0)
        monthKey = arrSplit(1)
        tempArr = dict(key)
        overallDPPM = 0: inspectedDPPM = 0
        If tempArr(0) > 0 Then overallDPPM = (tempArr(1) / tempArr(0)) * 1000000
        If tempArr(2) > 0 Then inspectedDPPM = (tempArr(3) / tempArr(2)) * 1000000
        outputArr(targetRow, 1) = supplier
        outputArr(targetRow, 2) = monthKey
        outputArr(targetRow, 3) = tempArr(0)
        outputArr(targetRow, 4) = tempArr(1)
        outputArr(targetRow, 5) = Format(overallDPPM, "0")
        outputArr(targetRow, 6) = tempArr(2)
        outputArr(targetRow, 7) = tempArr(3)
        outputArr(targetRow, 8) = Format(inspectedDPPM, "0")
        targetRow = targetRow + 1
    Next key
    BuildSupplierMonthlyOutputArray = outputArr
End Function

'=====================================================================================
' Private Sub: WriteSupplierMonthlySummary
' Purpose: Write Supplier-Monthly summary to worksheet and create table
'=====================================================================================
Private Sub WriteSupplierMonthlySummary(wsTarget As Worksheet, outputArr As Variant, tableName As String, procName As String)
    Dim dataRange As Range, tblTarget As ListObject
    wsTarget.Cells.Clear
    wsTarget.Cells(1, 1).value = "Supplier-Monthly DPPM Summary"
    wsTarget.Cells(1, 1).Font.Bold = True
    If UBound(outputArr, 1) > 1 Then
        Set dataRange = wsTarget.Range("A2").Resize(UBound(outputArr, 1), 8)
        dataRange.value = outputArr
    Else
        Set dataRange = wsTarget.Range("A2").Resize(1, 8)
        dataRange.value = Application.WorksheetFunction.Index(outputArr, 1, 0)
    End If
    On Error Resume Next
    wsTarget.ListObjects(tableName).Delete
    On Error GoTo 0
    Set tblTarget = wsTarget.ListObjects.Add(xlSrcRange, dataRange, , xlYes)
    tblTarget.Name = tableName
    tblTarget.TableStyle = Config.DEFAULT_TABLE_STYLE
    Utils.LogMessage "[" & procName & "] Created table '" & tableName & "' on sheet '" & wsTarget.Name & "'."
End Sub

