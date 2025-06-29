Attribute VB_Name = "GenerateDPPMTable"
'=====================================================================================
' Module: GenerateDPPMTable
' Purpose: Generate the DPPM output table from IQA Database and Wafer List
' Author: [Your Name]
' Date: [YYYY-MM-DD]
' Description:
'   - Aggregates IQA data by shipment and inspected date
'   - Handles wafer chip multiplication for specific suppliers
'   - Writes output to a formatted Excel table
'   - Triggers summary generation
'=====================================================================================
Option Explicit

' --- Module-level Variables ---
Private m_Config As Object

'=====================================================================================
' Public Sub: GenerateTable
' Purpose: Entry point to generate the DPPM output table and trigger summary
'=====================================================================================
Public Sub GenerateTable()
    Dim PROC_NAME As String: PROC_NAME = Config.PROC_GENERATE_TABLE
    Dim tblIQA As ListObject, wsTarget As Worksheet, wsWafer As Worksheet, tblTarget As ListObject, tblWafer As ListObject
    Dim dataArr As Variant, outputArr() As Variant
    Dim dataDict As Object
    Dim iqaWorkbook As Workbook
    Dim procStartTime As Double
    Dim lastRow As Long
    On Error GoTo GenericErrorHandler
    Utils.InitStatusBar PROC_NAME
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Utils.LogMessage "[" & PROC_NAME & "] Execution started."

    ' Load configuration
    Utils.UpdateStatusBarMessage "Loading configuration...", True
    Set m_Config = Utils.GetGlobalConfig()
    If m_Config Is Nothing Or m_Config.Count = 0 Then
        Utils.LogMessage "[" & PROC_NAME & "] Global configuration not loaded. Aborting.", True
        GoTo CleanUp
    End If

    ' Set the source table from the IQA Database
    Utils.UpdateStatusBarMessage "Setting up IQA Database...", True
    Set tblIQA = SetupIQADatabase(iqaWorkbook, m_Config)
    If tblIQA Is Nothing Then GoTo CleanUp
    Utils.UpdateStatusBarMessage "IQA Database setup complete.", stageComplete:=True

    ' Set the Wafer List sheet and table
    Utils.UpdateStatusBarMessage "Setting up Wafer List...", True
    Set wsWafer = Utils.GetSheet(Config.WAFER_LIST_SHEET_NAME)
    If wsWafer Is Nothing Then
        Utils.LogMessage "[" & PROC_NAME & "] Wafer List sheet '" & Config.WAFER_LIST_SHEET_NAME & "' not found!", True
        GoTo CleanUp
    End If
    On Error Resume Next
    Set tblWafer = wsWafer.ListObjects(Config.WAFER_LIST_TABLE_NAME)
    On Error GoTo GenericErrorHandler
    If tblWafer Is Nothing Then
        Utils.LogMessage "[" & PROC_NAME & "] Wafer List table '" & Config.WAFER_LIST_TABLE_NAME & "' not found on sheet '" & Config.WAFER_LIST_SHEET_NAME & "'!", True
        GoTo CleanUp
    End If
    Utils.UpdateStatusBarMessage "Wafer List setup.", stageComplete:=True

    ' Aggregate data
    dataArr = tblIQA.DataBodyRange.Value2
    lastRow = UBound(dataArr, 1)
    Set dataDict = AggregateDPPMData(dataArr, tblIQA, tblWafer, lastRow)
    Utils.LogMessage "[" & PROC_NAME & "] Data extraction and aggregation complete. " & dataDict.Count & " unique keys found."

    ' Build output array
    outputArr = BuildDPPMOutputArray(dataDict)

    ' Write to sheet and format
    Set wsTarget = tblIQA.Parent ' Output to same workbook as IQA source
    WriteDPPMTable wsTarget, outputArr, Config.DPPM_OUTPUT_TABLE_NAME
    FormatDPPMTable wsTarget, Config.DPPM_OUTPUT_TABLE_NAME

    ' Trigger summary generation
    Utils.LogMessage "[" & PROC_NAME & "] Calling GenerateSummary."
    Call GenerateDPPMSummary.GenerateSummary
    Utils.LogMessage "[" & PROC_NAME & "] Summary generation complete."
    MsgBox "DPPM table and summaries generated successfully!", vbInformation
    Utils.LogMessage "[" & PROC_NAME & "] Execution completed successfully."

CleanUp:
    If Not iqaWorkbook Is Nothing Then
        iqaWorkbook.Close SaveChanges:=False
        Set iqaWorkbook = Nothing
    End If
    Set wsTarget = Nothing
    Set wsWafer = Nothing
    Set tblTarget = Nothing
    Set tblWafer = Nothing
    Set dataDict = Nothing
    Set m_Config = Nothing
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Utils.ResetStatusBar PROC_NAME
    Exit Sub
GenericErrorHandler:
    Utils.LogMessage "[" & PROC_NAME & "] ERROR " & Err.Number & ": " & Err.Description & " (Line: " & Erl & ")", True
    Utils.ResetStatusBar PROC_NAME, True, Err.Description
    MsgBox "An error occurred in " & PROC_NAME & ": " & Err.Description & vbCrLf & "Please check the ExecutionLog.txt for details.", vbCritical
    Resume CleanUp
End Sub

'=====================================================================================
' Private Function: AggregateDPPMData
' Purpose: Aggregate IQA data by shipment and inspected date, handling wafer logic
'=====================================================================================
Private Function AggregateDPPMData(dataArr As Variant, tblIQA As ListObject, tblWafer As ListObject, lastRow As Long) As Object
    Dim dataDict As Object
    Dim i As Long
    Dim key As String, tempArray As Variant
    Dim shipmentDate As Variant, inspectedDate As Variant, supplierName As String, partNumber As String, inspectedBy As String
    Dim quantityIn As Double, rejectQuantity As Double, chipsPerWaferCount As Double
    Dim colIdxShipDate As Long, colIdxInspDate As Long, colIdxSupplier As Long, colIdxPartNum As Long, colIdxInspBy As Long, colIdxQtyIn As Long, colIdxRejQty As Long
    Set dataDict = CreateObject("Scripting.Dictionary")
    colIdxShipDate = tblIQA.ListColumns(Config.IQA_COL_SHIP_DATE).Index
    colIdxInspDate = tblIQA.ListColumns(Config.IQA_COL_INSPECTED_BY).Index
    colIdxSupplier = tblIQA.ListColumns(Config.IQA_COL_SUPPLIER).Index
    colIdxPartNum = tblIQA.ListColumns(Config.IQA_COL_PART_NUM).Index
    colIdxInspBy = tblIQA.ListColumns(Config.IQA_COL_INSPECTED_BY).Index
    colIdxQtyIn = tblIQA.ListColumns(Config.IQA_COL_QUANTITY_IN).Index
    colIdxRejQty = tblIQA.ListColumns(Config.IQA_COL_TOTAL_REJECT_QUANTITY).Index
    Dim procStartTime As Double: procStartTime = Timer
    For i = 1 To lastRow
        If i Mod Utils.STATUS_BAR_RECORD_UPDATE_INTERVAL = 0 Or Timer - Utils.g_lngLastStatusBarUpdateTime > Utils.STATUS_BAR_UPDATE_INTERVAL_SECONDS Then
            Utils.UpdateStatusBarProgress "Aggregating DPPM Data", i, lastRow, procStartTime
        End If

        If IsNumeric(dataArr(i, colIdxShipDate)) And (dataArr(i, colIdxShipDate) > 0) Then
            ' Convert Serial to Date
            shipmentDate = Format(CDate(dataArr(i, colIdxShipDate)), "yyyy-MM-dd")
            If Not IsDate(shipmentDate) Then
                Utils.LogMessage "[" & Config.PROC_GENERATE_TABLE & "] Invalid shipment date at row " & i & ". Skipping row."
                GoTo NextRow
            End If
        End If

        If IsNumeric(dataArr(i, colIdxInspDate)) And (dataArr(i, colIdxInspDate) > 0) Then
            ' Convert Serial to Date
            inspectedDate = Format(CDate(dataArr(i, colIdxInspDate)), "yyyy-MM-dd")
            If Not IsDate(inspectedDate) Then
                inspectedDate = shipmentDate ' If inspection date is invalid, use shipment date
                Utils.LogMessage "[" & Config.PROC_GENERATE_TABLE & "] No inspection date found, using shipment date for row " & i & "."
            End If
        Else
            inspectedDate = shipmentDate ' If no inspection date, use shipment date
            Utils.LogMessage "[" & Config.PROC_GENERATE_TABLE & "] No inspection date found, using shipment date for row " & i & "."
        End If

        supplierName = Trim(CStr(dataArr(i, colIdxSupplier)))
        partNumber = Trim(CStr(dataArr(i, colIdxPartNum)))
        inspectedBy = Trim(CStr(dataArr(i, colIdxInspBy)))
        quantityIn = 0
        If IsNumeric(dataArr(i, colIdxQtyIn)) Then quantityIn = CDbl(dataArr(i, colIdxQtyIn))
        rejectQuantity = 0
        If IsNumeric(dataArr(i, colIdxRejQty)) Then rejectQuantity = CDbl(dataArr(i, colIdxRejQty))
        chipsPerWaferCount = 0
        If supplierName = "EXCELITAS CANADA INC." Then
            chipsPerWaferCount = GetChipsPerWafer(tblWafer, partNumber)
            If chipsPerWaferCount > 0 Then quantityIn = quantityIn * chipsPerWaferCount
        End If
        If shipmentDate = "" Or supplierName = "" Or partNumber = "" Then GoTo NextRow
        key = shipmentDate & "|" & supplierName & "|" & partNumber
        If Not dataDict.exists(key) Then dataDict.Add key, Array(shipmentDate, supplierName, partNumber, inspectedBy, 0, 0, 0, 0)
        tempArray = dataDict(key)
        tempArray(4) = tempArray(4) + quantityIn
        tempArray(5) = tempArray(5) + rejectQuantity
        dataDict(key) = tempArray
        If IsDate(inspectedDate) Then
            key = inspectedDate & "|" & supplierName & "|" & partNumber
            If Not dataDict.exists(key) Then dataDict.Add key, Array(inspectedDate, supplierName, partNumber, inspectedBy, 0, 0, 0, 0)
            tempArray = dataDict(key)
            tempArray(6) = tempArray(6) + quantityIn
            tempArray(7) = tempArray(7) + rejectQuantity
            dataDict(key) = tempArray
        End If
NextRow:
    Next i
    Set AggregateDPPMData = dataDict
End Function

'=====================================================================================
' Private Function: BuildDPPMOutputArray
' Purpose: Build output array for DPPM table from dictionary
'=====================================================================================
Private Function BuildDPPMOutputArray(dataDict As Object) As Variant
    Dim outputArr() As Variant, tempArray As Variant, key As Variant
    Dim overallDPPM As Double, inspectedDPPM As Double
    Dim targetRow As Long
    ReDim outputArr(1 To dataDict.Count + 1, 1 To 10)
    outputArr(1, 1) = Config.DPPM_COL_DATE
    outputArr(1, 2) = Config.DPPM_COL_SUPPLIER
    outputArr(1, 3) = Config.DPPM_COL_PART_NUM
    outputArr(1, 4) = Config.DPPM_COL_INSPECTED_BY
    outputArr(1, 5) = Config.DPPM_COL_OVERALL_QTY
    outputArr(1, 6) = Config.DPPM_COL_OVERALL_REJECT
    outputArr(1, 7) = Config.DPPM_COL_OVERALL_DPPM
    outputArr(1, 8) = Config.DPPM_COL_INSPECTED_QTY
    outputArr(1, 9) = Config.DPPM_COL_INSPECTED_REJECT
    outputArr(1, 10) = Config.DPPM_COL_INSPECTED_DPPM
    targetRow = 2
    For Each key In dataDict.keys
        tempArray = dataDict(key)
        overallDPPM = 0: inspectedDPPM = 0
        If tempArray(4) > 0 Then overallDPPM = (tempArray(5) / tempArray(4)) * 1000000
        If tempArray(6) > 0 Then inspectedDPPM = (tempArray(7) / tempArray(6)) * 1000000
        outputArr(targetRow, 1) = tempArray(0)
        outputArr(targetRow, 2) = tempArray(1)
        outputArr(targetRow, 3) = tempArray(2)
        outputArr(targetRow, 4) = tempArray(3)
        outputArr(targetRow, 5) = tempArray(4)
        outputArr(targetRow, 6) = tempArray(5)
        outputArr(targetRow, 7) = Format(overallDPPM, "0")
        outputArr(targetRow, 8) = tempArray(6)
        outputArr(targetRow, 9) = tempArray(7)
        outputArr(targetRow, 10) = Format(inspectedDPPM, "0")
        targetRow = targetRow + 1
    Next key
    BuildDPPMOutputArray = outputArr
End Function

'=====================================================================================
' Private Sub: WriteDPPMTable
' Purpose: Write DPPM output array to worksheet and create table
'=====================================================================================
Private Sub WriteDPPMTable(wsTarget As Worksheet, outputArr As Variant, tableName As String)
    Dim PROC_NAME As String: PROC_NAME = Config.PROC_GENERATE_TABLE_WRITE
    Dim dataRange As Range, tblTarget As ListObject
    wsTarget.Cells.ClearContents
    If UBound(outputArr, 1) > 1 Then
        Set dataRange = wsTarget.Range("A1").Resize(UBound(outputArr, 1), 10)
        dataRange.Value = outputArr
    Else
        Set dataRange = wsTarget.Range("A1").Resize(1, 10)
        dataRange.Value = Application.WorksheetFunction.Index(outputArr, 1, 0)
    End If
    On Error Resume Next
    wsTarget.ListObjects(tableName).Delete
    On Error GoTo 0
    Set tblTarget = wsTarget.ListObjects.Add(xlSrcRange, dataRange, , xlYes)
    tblTarget.Name = tableName
    tblTarget.TableStyle = Config.DEFAULT_TABLE_STYLE
    Utils.LogMessage "[" & PROC_NAME & "] DPPM data written to table '" & tableName & "'."
    Utils.UpdateStatusBarMessage "DPPM data written.", stageComplete:=True
End Sub

'=====================================================================================
' Private Sub: FormatDPPMTable
' Purpose: Sort and format the DPPM table
'=====================================================================================
Private Sub FormatDPPMTable(wsTarget As Worksheet, tableName As String)
    Dim PROC_NAME As String: PROC_NAME = Config.PROC_GENERATE_TABLE_FORMAT
    Dim tblTarget As ListObject
    Set tblTarget = wsTarget.ListObjects(tableName)
    If tblTarget.ListRows.Count > 0 Then
        Utils.UpdateStatusBarMessage "Sorting DPPM table...", True
        With tblTarget.Sort
            .SortFields.Clear
            .SortFields.Add Key:=tblTarget.ListColumns(Config.DPPM_COL_DATE).Range, SortOn:=xlSortOnValues, Order:=xlAscending
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
        Utils.LogMessage "[" & PROC_NAME & "] Sorted table by " & Config.DPPM_COL_DATE & "."
        Utils.UpdateStatusBarMessage "DPPM table sorted.", stageComplete:=True
    End If
    tblTarget.Range.Columns.AutoFit
    With tblTarget.Range
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    Utils.LogMessage "[" & PROC_NAME & "] Applied formatting to table."
    Utils.UpdateStatusBarMessage "DPPM table formatted.", stageComplete:=True
End Sub

'=====================================================================================
' Private Function: GetChipsPerWafer
' Purpose: Get chips per wafer for a given part number from wafer table
'=====================================================================================
Private Function GetChipsPerWafer(tblWafer As ListObject, partNumber As String) As Double
    Dim waferPartArr As Variant
    Dim waferRow As Long
    Dim idxPartNum As Long, idxChips As Long
    GetChipsPerWafer = 0
    If tblWafer Is Nothing Then Exit Function
    On Error Resume Next
    idxPartNum = tblWafer.ListColumns(Config.WAFER_LIST_COL_PART_NUM).Index
    idxChips = tblWafer.ListColumns(Config.WAFER_LIST_COL_CHIPS_PER_WAFER).Index
    waferPartArr = tblWafer.DataBodyRange.Value
    For waferRow = 1 To UBound(waferPartArr, 1)
        If Trim(CStr(waferPartArr(waferRow, idxPartNum))) = partNumber Then
            GetChipsPerWafer = waferPartArr(waferRow, idxChips)
            Exit For
        End If
    Next waferRow
    On Error GoTo 0
End Function

'=====================================================================================
' Private Function: SetupIQADatabase
' Purpose: Open IQA Database and return the source table
'=====================================================================================
Private Function SetupIQADatabase(ByRef iqaWorkbook As Workbook, ByVal cfg As Object) As ListObject
    Dim PROC_NAME As String: PROC_NAME = Config.PROC_LOAD_IQA_DB
    Dim iqaSource As Worksheet
    Dim iqaDatabasePath As String
    On Error GoTo ErrorHandler
    iqaDatabasePath = ""
    If Not cfg Is Nothing And cfg.exists(Config.CONFIG_KEY_IQA_DB_PATH) Then
        iqaDatabasePath = cfg(Config.CONFIG_KEY_IQA_DB_PATH)
    End If
    If iqaDatabasePath = "" Then
        Utils.LogMessage "[" & PROC_NAME & "] IQA Database path key '" & Config.CONFIG_KEY_IQA_DB_PATH & "' not found or empty in configuration!", True
        Exit Function
    End If
    Set iqaWorkbook = Workbooks.Open(Filename:=iqaDatabasePath, ReadOnly:=True, IgnoreReadOnlyRecommended:=True, UpdateLinks:=False)
    If iqaWorkbook Is Nothing Then
        Utils.LogMessage "[" & PROC_NAME & "] Failed to open IQA Database at: " & iqaDatabasePath, True
        Exit Function
    End If
    Set iqaSource = iqaWorkbook.Sheets(Config.IQA_SHEET_NAME)
    If iqaSource Is Nothing Then
        Utils.LogMessage "[" & PROC_NAME & "] Sheet '" & Config.IQA_SHEET_NAME & "' not found in the IQA workbook!", True
        MsgBox "Sheet '" & Config.IQA_SHEET_NAME & "' not found in the IQA workbook!", vbExclamation
        Exit Function
    End If
    Dim tblIQA As ListObject
    Set tblIQA = Nothing
    Set tblIQA = iqaSource.ListObjects(Config.IQA_TABLE_NAME)
    If tblIQA Is Nothing Then
        Utils.LogMessage "[" & PROC_NAME & "] Table '" & Config.IQA_TABLE_NAME & "' not found in the IQA sheet '" & Config.IQA_SHEET_NAME & "'.", True
        MsgBox "Table '" & Config.IQA_TABLE_NAME & "' not found in the IQA sheet '" & Config.IQA_SHEET_NAME & "'.", vbExclamation
        Exit Function
    End If
    If tblIQA.AutoFilter.FilterMode Then tblIQA.AutoFilter.ShowAllData
    Utils.LogMessage "[" & PROC_NAME & "] IQA Database setup successfully from: " & iqaDatabasePath, False
    Set SetupIQADatabase = tblIQA
    Exit Function
ErrorHandler:
    Utils.LogMessage "[" & PROC_NAME & "] Error setting up IQA Database: " & Err.Description, True
    Application.EnableEvents = True
    Set SetupIQADatabase = Nothing
End Function

'=====================================================================================
' [Unit Test Stubs]
'=====================================================================================
' TODO: Add unit tests for AggregateDPPMData and BuildDPPMOutputArray
' Example:
' 'Public Sub Test_AggregateDPPMData()
' '    ' Arrange: Create mock data array and tables
' '    ' Act: Call AggregateDPPMData
' '    ' Assert: Check dictionary contents
' 'End Sub
