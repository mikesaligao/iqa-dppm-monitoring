Attribute VB_Name = "GenerateDPPMTable"

Option Explicit
' Declare a module-level variable To store configuration
Private m_Config As Object

' Module-level variables for source IQA column indices (populated from config)
Private m_SourceShipmentDateColName As String
Private m_InspectedDateColName As String
Private m_SupplierNameColName As String
Private m_PartNumberColName As String
Private m_InspectedByColName As String
Private m_QuantityInColName As String
Private m_RejectQuantityColName As String

Private m_SourceShipmentDateColIdx As Long
Private m_InspectedDateColIdx As Long
Private m_SupplierNameColIdx As Long
Private m_PartNumberColIdx As Long
Private m_InspectedByColIdx As Long
Private m_QuantityInColIdx As Long
Private m_RejectQuantityColIdx As Long

Private Const PROC_NAME As String = "GenerateDPPMTable"

Public Sub GenerateTable()
    Dim wsSource As Worksheet, wsTarget As Worksheet, wsWafer As Worksheet, tblTarget As ListObject, tblWafer As ListObject
    Dim lastRow As Long, targetRow As Long, i As Long
    Dim key As Variant, tempArray As Variant, dataArr As Variant, outputArr() As Variant
    Dim shipmentDate As String, inspectedDateVal As Variant, supplierName As String, partNumber As String, inspectedBy As String
    Dim quantityIn As Double, rejectQuantity As Double, chipsPerWaferCount As Double
    Dim overallQuantity As Double, overallRejects As Double, overallDPPM As Double
    Dim inspectedQuantity As Double, inspectedRejects As Double, inspectedDPPM As Double
    Dim dataDict As Object
    Dim iqaWorkbook As Workbook
    Dim procStartTime As Double

    On Error GoTo GenericErrorHandler
    Utils.InitStatusBar PROC_NAME

        ' Disable screen updating For performance
        Application.ScreenUpdating = False
        Application.Calculation = xlCalculationManual
        Application.EnableEvents = False
        Utils.LogMessage "[" & PROC_NAME & "] Execution started."

        ' Load configuration from the Config sheet
        Utils.UpdateStatusBarMessage "Loading configuration...", True
        Set m_Config = Utils.GetGlobalConfig()
        If m_Config Is Nothing Or m_Config.Count = 0 Then
            Utils.LogMessage "[" & PROC_NAME & "] Global configuration not loaded. Aborting.", True
            GoTo CleanUp
        End If
        If Not LoadTableModuleConfig(m_Config) Then GoTo CleanUp ' Error logged in LoadTableModuleConfig
        Utils.UpdateStatusBarMessage "Configuration loaded.", stageComplete:=True

        ' Set the source sheet from the current workbook
        Utils.UpdateStatusBarMessage "Setting up IQA Database...", True
        Set wsSource = SetupIQADatabase(iqaWorkbook, m_Config)
        If wsSource Is Nothing Then GoTo CleanUp ' Error logged in SetupIQADatabase or LoadTableModuleConfig
        Utils.UpdateStatusBarMessage "IQA Database setup complete.", stageComplete:=True

        ' Get or create the target sheet and table
        Utils.UpdateStatusBarMessage "Setting up target DPPM table...", True
        Set wsTarget = Utils.GetSheet(Config.DPPM_OUTPUT_SHEET_NAME)
        If wsTarget Is Nothing Then
            Set wsTarget = ThisWorkbook.Sheets.Add
            wsTarget.Name = Config.DPPM_OUTPUT_SHEET_NAME
            Utils.LogMessage "[" & PROC_NAME & "] Created target sheet: '" & Config.DPPM_OUTPUT_SHEET_NAME & "'"
        End If

        ' Clear existing table if it exists, or clear sheet for new table
        On Error Resume Next
        If Not tblTarget Is Nothing Then
            If Not tblTarget.DataBodyRange Is Nothing Then
                tblTarget.DataBodyRange.Delete
            End If
            Utils.LogMessage "[" & PROC_NAME & "] Cleared data from existing table: '" & Config.DPPM_OUTPUT_TABLE_NAME & "'"
        Else
            wsTarget.Cells.ClearContents
            Utils.LogMessage "[" & PROC_NAME & "] Cleared sheet for new table: '" & Config.DPPM_OUTPUT_SHEET_NAME & "'"
        End If
        On Error GoTo GenericErrorHandler
        Utils.UpdateStatusBarMessage "Target DPPM table setup.", stageComplete:=True

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

        ' Initialize dictionary For aggregation
        Utils.LogMessage "[" & PROC_NAME & "] Initializing dictionary for data aggregation."
        Set dataDict = CreateObject("Scripting.Dictionary")

        ' Find the last row in the source sheet
        lastRow = wsSource.Cells(wsSource.Rows.Count, "B").End(xlUp).Row ' Assuming Col B (Supplier) is reliable for last row
        Utils.LogMessage "[" & PROC_NAME & "] Last row in source IQA sheet: " & lastRow

        ' Read the entire relevant range from source IQA sheet into an array
        ' Determine the maximum column index needed from the source sheet
        Dim maxColIdx As Long
        maxColIdx = Application.WorksheetFunction.Max(m_SourceShipmentDateColIdx, m_InspectedDateColIdx, m_SupplierNameColIdx, _
                                                    m_PartNumberColIdx, m_InspectedByColIdx, m_QuantityInColIdx, m_RejectQuantityColIdx)

        If maxColIdx = 0 Then
            Utils.LogMessage "[" & PROC_NAME & "] One or more source column indices are invalid (0). Aborting.", True
            GoTo CleanUp
        End If

        If lastRow >= 2 Then ' Ensure there's data beyond headers
            ' Read up to the maximum column index required
            Utils.LogMessage "[" & PROC_NAME & "] Reading data from source IQA sheet range: A2:" & Cells(lastRow, maxColIdx).Address(False, False), False
            dataArr = wsSource.Range(wsSource.Cells(2, 1), wsSource.Cells(lastRow, maxColIdx)).Value ' Using .Value as dates are involved and .Value2 might return numbers for dates
        Else
            Utils.LogMessage "[" & PROC_NAME & "] No data found in source IQA sheet.", True
            GoTo OutputHeadersOnly ' Proceed to output headers if no data
        End If

        ' Loop through each row in the source data array
        Utils.LogMessage "[" & PROC_NAME & "] Starting data extraction and aggregation."
        procStartTime = Timer
        For i = 1 To UBound(dataArr, 1) ' dataArr is 1-based
            If i Mod Utils.STATUS_BAR_RECORD_UPDATE_INTERVAL = 0 Or Timer - Utils.g_lngLastStatusBarUpdateTime > Utils.STATUS_BAR_UPDATE_INTERVAL_SECONDS Then
                Utils.UpdateStatusBarProgress "Aggregating DPPM Data", i, UBound(dataArr, 1), procStartTime
            End If

            shipmentDate = Format(dataArr(i, m_SourceShipmentDateColIdx), "yyyy-MM-dd")
            inspectedDateVal = dataArr(i, m_InspectedDateColIdx) ' Keep as variant for IsDate check
            supplierName = Trim(CStr(dataArr(i, m_SupplierNameColIdx)))
            partNumber = Trim(CStr(dataArr(i, m_PartNumberColIdx)))
            inspectedBy = Trim(CStr(dataArr(i, m_InspectedByColIdx)))

            quantityIn = 0
            If IsNumeric(dataArr(i, m_QuantityInColIdx)) Then quantityIn = CDbl(dataArr(i, m_QuantityInColIdx))

            rejectQuantity = 0
            If IsNumeric(dataArr(i, m_RejectQuantityColIdx)) Then rejectQuantity = CDbl(dataArr(i, m_RejectQuantityColIdx))

            'Utils.LogMessage "[" & PROC_NAME & "] Row " & i + 1 & ": ShipDate=" & shipmentDate & ", Supp=" & supplierName & ", PN=" & partNumber & ", QtyIn=" & quantityIn & ", RejQty=" & rejectQuantity, False

            chipsPerWaferCount = 0
            If supplierName = "EXCELITAS CANADA INC." Then
                'Utils.LogMessage "[" & PROC_NAME & "] Row " & i + 1 & ": ECI supplier. Looking up chips/wafer for PN: " & partNumber, False
                On Error Resume Next
                chipsPerWaferCount = Application.WorksheetFunction.VLookup(partNumber, tblWafer.Range, Utils.GetColumnIndexByName(tblWafer, Config.WAFER_LIST_COL_CHIPS), False)
                If Err.Number <> 0 Then
                    chipsPerWaferCount = 0 ' Not found or error
                    'Utils.LogMessage "[" & PROC_NAME & "] Row " & i + 1 & ": VLookup for PN '" & partNumber & "' in Wafer List failed. Error: " & Err.Description, False
                    Err.Clear
                End If
                On Error GoTo GenericErrorHandler

                If chipsPerWaferCount <= 0 Then
                    'Utils.LogMessage "[" & PROC_NAME & "] Row " & i + 1 & ": Chips/wafer for PN '" & partNumber & "' not found or invalid (" & chipsPerWaferCount & "). Skipping adjustment.", False
                    ' Decide if row should be skipped or processed with original quantityIn
                    ' Current logic implies original quantityIn is used if lookup fails or is <=0, then adjusted if >0
                    ' For safety, let's ensure it's not an error causing skip:
                    ' GoTo NextRow ' If this is critical
                Else
                    'Utils.LogMessage "[" & PROC_NAME & "] Row " & i + 1 & ": Chips/wafer found: " & chipsPerWaferCount & ". Original QtyIn: " & quantityIn, False
                    quantityIn = quantityIn * chipsPerWaferCount
                    'Utils.LogMessage "[" & PROC_NAME & "] Row " & i + 1 & ": Adjusted QtyIn: " & quantityIn, False
                End If
            End If

            If shipmentDate = "" Or supplierName = "" Or partNumber = "" Then
                'Utils.LogMessage "[" & PROC_NAME & "] Row " & i + 1 & ": Missing key data (ShipDate, Supplier, or PartNumber). Skipping.", False
                GoTo NextRow
            End If

            ' Create a unique key for grouping by shipment date
            key = shipmentDate & "|" & supplierName & "|" & partNumber
            If Not dataDict.exists(key) Then
                ' Array: 0:Date, 1:Supplier, 2:PartNum, 3:InspectedBy,
                '        4:OverallQty, 5:OverallRejects, 6:InspectedQty, 7:InspectedRejects
                dataDict.Add key, Array(shipmentDate, supplierName, partNumber, inspectedBy, 0, 0, 0, 0)
            End If
            tempArray = dataDict(key)
            tempArray(4) = tempArray(4) + quantityIn     ' Overall Quantity Received
            tempArray(5) = tempArray(5) + rejectQuantity ' Overall Units Reject
            dataDict(key) = tempArray

            ' Aggregate by inspected date if valid
            If IsDate(inspectedDateVal) Then
                Dim formattedInspectedDate As String
                formattedInspectedDate = Format(CDate(inspectedDateVal), "yyyy-MM-dd") ' Ensure it's treated as date before formatting
                key = formattedInspectedDate & "|" & supplierName & "|" & partNumber ' Use the same key structure

                If Not dataDict.exists(key) Then
                    dataDict.Add key, Array(formattedInspectedDate, supplierName, partNumber, inspectedBy, 0, 0, 0, 0)
                End If
                tempArray = dataDict(key)
                tempArray(6) = tempArray(6) + quantityIn     ' Inspected Quantity Received
                tempArray(7) = tempArray(7) + rejectQuantity ' Inspected Units Reject
                dataDict(key) = tempArray
            Else
                'Utils.LogMessage "[" & PROC_NAME & "] Row " & i + 1 & ": Invalid Inspected Date (" & CStr(inspectedDateVal) & "). Skipping inspected aggregation for this entry.", False
            End If

NextRow:
        Next i
        If UBound(dataArr, 1) > 0 Then Utils.UpdateStatusBarProgress "Aggregating DPPM Data", UBound(dataArr, 1), UBound(dataArr, 1), procStartTime ' Final update
        Utils.LogMessage "[" & PROC_NAME & "] Data extraction and aggregation complete. " & dataDict.Count & " unique keys found."

OutputHeadersOnly:
        ' Pre-size the output array: [1 To dict.Count + 1 (for headers), 1 To 10 columns]
        ReDim outputArr(1 To dataDict.Count + 1, 1 To 10)

        ' Write headers to the output array
        Utils.LogMessage "[" & PROC_NAME & "] Preparing headers for the target table."
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

        ' Write aggregated data to the output array
        Utils.LogMessage "[" & PROC_NAME & "] Writing aggregated data to array."
        targetRow = 2 ' Start data from the second row of the array
        For Each key In dataDict.keys
            tempArray = dataDict(key)
            shipmentDate = tempArray(0) ' This is the date (either shipment or inspected)
            supplierName = tempArray(1)
            partNumber = tempArray(2)
            inspectedBy = tempArray(3)
            overallQuantity = tempArray(4)
            overallRejects = tempArray(5)
            inspectedQuantity = tempArray(6)
            inspectedRejects = tempArray(7)

            overallDPPM = 0
            If overallQuantity > 0 Then overallDPPM = (overallRejects / overallQuantity) * 1000000

            inspectedDPPM = 0
            If inspectedQuantity > 0 Then inspectedDPPM = (inspectedRejects / inspectedQuantity) * 1000000

            outputArr(targetRow, 1) = shipmentDate
            outputArr(targetRow, 2) = supplierName
            outputArr(targetRow, 3) = partNumber
            outputArr(targetRow, 4) = inspectedBy
            outputArr(targetRow, 5) = overallQuantity
            outputArr(targetRow, 6) = overallRejects
            outputArr(targetRow, 7) = Format(overallDPPM, "0")
            outputArr(targetRow, 8) = inspectedQuantity
            outputArr(targetRow, 9) = inspectedRejects
            outputArr(targetRow, 10) = Format(inspectedDPPM, "0")

            targetRow = targetRow + 1
        Next key

        ' Write all data from array to the sheet and create/update table
        Utils.UpdateStatusBarMessage "Writing data to DPPM table...", True
        If wsTarget.ListObjects.Count > 0 Then ' Check if any table exists
            On Error Resume Next
            wsTarget.ListObjects(Config.DPPM_OUTPUT_TABLE_NAME).Delete
            On Error GoTo GenericErrorHandler
        End If
        wsTarget.Cells.ClearContents ' Clear sheet before writing new data/table

        Dim dataRange As Range
        If dataDict.Count > 0 Then
            Set dataRange = wsTarget.Range("A1").Resize(dataDict.Count + 1, 10)
            dataRange.Value = outputArr
            Set tblTarget = wsTarget.ListObjects.Add(xlSrcRange, dataRange, , xlYes)
        Else ' Only headers if no data
            Set dataRange = wsTarget.Range("A1").Resize(1, 10)
            dataRange.Value = Application.WorksheetFunction.Index(outputArr, 1, 0) ' Get first row (headers)
            Set tblTarget = wsTarget.ListObjects.Add(xlSrcRange, dataRange, , xlYes)
        End If

        tblTarget.Name = Config.DPPM_OUTPUT_TABLE_NAME
        tblTarget.TableStyle = Config.DEFAULT_TABLE_STYLE
        Utils.LogMessage "[" & PROC_NAME & "] DPPM data written to table '" & Config.DPPM_OUTPUT_TABLE_NAME & "'."
        Utils.UpdateStatusBarMessage "DPPM data written.", stageComplete:=True

        ' Sort the data by Date (first column)
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

        ' Auto-fit columns and center align
        Utils.UpdateStatusBarMessage "Formatting DPPM table...", True
        tblTarget.Range.Columns.AutoFit
        With tblTarget.Range
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        Utils.LogMessage "[" & PROC_NAME & "] Applied formatting to table."
        Utils.UpdateStatusBarMessage "DPPM table formatted.", stageComplete:=True

        ' Database info is updated, Update the Summary sheet
        Utils.LogMessage "[" & PROC_NAME & "] Calling GenerateSummary."
        Call GenerateDPPMSummary.GenerateSummary ' Explicitly call from module
        Utils.LogMessage "[" & PROC_NAME & "] Summary generation complete."

        MsgBox "DPPM table and summaries generated successfully!", vbInformation
        Utils.LogMessage "[" & PROC_NAME & "] Execution completed successfully."

Cleanup:
        If Not iqaWorkbook Is Nothing Then
            iqaWorkbook.Close SaveChanges:=False
            Set iqaWorkbook = Nothing
        End If
        Set wsSource = Nothing
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

Private Function LoadTableModuleConfig(ByVal cfg As Object) As Boolean
    ' Loads and validates module-specific configuration from the global config object (cfg)
    ' Populates module-level variables for IQA source column NAMES.
    ' Column indices will be determined later once the source sheet/table is known.
    Dim funcName As String: funcName = "LoadTableModuleConfig"
    Dim missingKeys As String
    Dim keyName As Variant
    Dim requiredColNameKeys As Variant

    On Error GoTo ErrorHandler
    LoadTableModuleConfig = False ' Default to failure

    If cfg Is Nothing Or cfg.Count = 0 Then
        Utils.LogMessage "[" & PROC_NAME & "] " & funcName & ": Global configuration object is empty or not provided.", True
        Exit Function
    End If

    ' Define the configuration keys that provide the COLUMN NAMES for IQA source data
    requiredColNameKeys = Array( _
        Config.CONFIG_KEY_IQA_SRC_SHIP_DATE_COLNAME, _
        Config.CONFIG_KEY_IQA_SRC_INSP_DATE_COLNAME, _
        Config.CONFIG_KEY_IQA_SRC_SUPPLIER_COLNAME, _
        Config.CONFIG_KEY_IQA_SRC_PARTNUM_COLNAME, _
        Config.CONFIG_KEY_IQA_SRC_INSP_BY_COLNAME, _
        Config.CONFIG_KEY_IQA_SRC_QTY_IN_COLNAME, _
        Config.CONFIG_KEY_IQA_SRC_REJ_QTY_COLNAME _
    )

    ' Validate that these keys exist in the global config
    For Each keyName In requiredColNameKeys
        If Not cfg.exists(CStr(keyName)) Or Trim(CStr(cfg(CStr(keyName)))) = "" Then
            missingKeys = missingKeys & vbCrLf & " - " & CStr(keyName)
        End If
    Next keyName

    If Len(missingKeys) > 0 Then
        Utils.LogMessage "[" & PROC_NAME & "] " & funcName & ": Missing or empty configuration for the following keys:" & missingKeys, True
        MsgBox "Configuration error in " & PROC_NAME & ":" & vbCrLf & _
               "The following required settings for IQA source column names are missing or empty in the 'Config' sheet:" & _
               missingKeys & vbCrLf & vbCrLf & "Please check the 'Config' sheet and the ExecutionLog.txt for details.", vbCritical
        Exit Function
    End If

    ' Assign column names from config to module-level variables
    m_SourceShipmentDateColName = Trim(CStr(cfg(Config.CONFIG_KEY_IQA_SRC_SHIP_DATE_COLNAME)))
    m_InspectedDateColName = Trim(CStr(cfg(Config.CONFIG_KEY_IQA_SRC_INSP_DATE_COLNAME)))
    m_SupplierNameColName = Trim(CStr(cfg(Config.CONFIG_KEY_IQA_SRC_SUPPLIER_COLNAME)))
    m_PartNumberColName = Trim(CStr(cfg(Config.CONFIG_KEY_IQA_SRC_PARTNUM_COLNAME)))
    m_InspectedByColName = Trim(CStr(cfg(Config.CONFIG_KEY_IQA_SRC_INSP_BY_COLNAME)))
    m_QuantityInColName = Trim(CStr(cfg(Config.CONFIG_KEY_IQA_SRC_QTY_IN_COLNAME)))
    m_RejectQuantityColName = Trim(CStr(cfg(Config.CONFIG_KEY_IQA_SRC_REJ_QTY_COLNAME)))

    Utils.LogMessage "[" & PROC_NAME & "] " & funcName & ": Module column name configuration loaded successfully.", False
    LoadTableModuleConfig = True
    Exit Function

ErrorHandler:
    Utils.LogMessage "[" & PROC_NAME & "] " & funcName & ": ERROR " & Err.Number & " - " & Err.Description, True
    LoadTableModuleConfig = False
End Function

Private Function FindAndSetColumnIndex(ByVal headerRow As Range, ByVal colNameConfig As String, ByRef outIndexVar As Long, ByVal wsNameForLog As String) As Boolean
    ' Helper function to find a column by name in a header row and set its index.
    Dim foundCell As Range
    Dim funcName As String: funcName = "FindAndSetColumnIndex"
    On Error Resume Next ' Keep local error handling for Find
    Set foundCell = headerRow.Find(What:=colNameConfig, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
    On Error GoTo 0
    If Not foundCell Is Nothing Then
        outIndexVar = foundCell.Column
        FindAndSetColumnIndex = True
    Else
        Utils.LogMessage "[" & PROC_NAME & "] " & funcName & ": Column '" & colNameConfig & "' not found in '" & wsNameForLog & "' header row.", True
        FindAndSetColumnIndex = False
    End If
End Function
Private Function GetColumnIndicesFromSource(wsSource As Worksheet) As Boolean
    ' Determines column indices based on names stored in module-level variables.
    ' This should be called AFTER wsSource is set.
    Dim funcName As String: funcName = "GetColumnIndicesFromSource"
    Dim headerRow As Range
    Dim allIndicesFound As Boolean

    On Error GoTo ErrorHandler
    GetColumnIndicesFromSource = False ' Default to failure
    allIndicesFound = True

    If wsSource Is Nothing Then
        Utils.LogMessage "[" & PROC_NAME & "] " & funcName & ": Source worksheet is not set.", True
        Exit Function
    End If

    Set headerRow = wsSource.Rows(1) ' Assuming headers are in row 1

    ' Get indices for all required columns
    ' Use the new helper function
    If Not FindAndSetColumnIndex(headerRow, m_SourceShipmentDateColName, m_SourceShipmentDateColIdx, wsSource.Name) Then allIndicesFound = False
    If Not FindAndSetColumnIndex(headerRow, m_InspectedDateColName, m_InspectedDateColIdx, wsSource.Name) Then allIndicesFound = False
    If Not FindAndSetColumnIndex(headerRow, m_SupplierNameColName, m_SupplierNameColIdx, wsSource.Name) Then allIndicesFound = False
    If Not FindAndSetColumnIndex(headerRow, m_PartNumberColName, m_PartNumberColIdx, wsSource.Name) Then allIndicesFound = False
    If Not FindAndSetColumnIndex(headerRow, m_InspectedByColName, m_InspectedByColIdx, wsSource.Name) Then allIndicesFound = False
    If Not FindAndSetColumnIndex(headerRow, m_QuantityInColName, m_QuantityInColIdx, wsSource.Name) Then allIndicesFound = False
    If Not FindAndSetColumnIndex(headerRow, m_RejectQuantityColName, m_RejectQuantityColIdx, wsSource.Name) Then allIndicesFound = False

    If allIndicesFound Then
        Utils.LogMessage "[" & PROC_NAME & "] " & funcName & ": All source column indices determined successfully.", False
        GetColumnIndicesFromSource = True
    Else
        MsgBox "One or more required columns were not found in the IQA Database sheet ('" & wsSource.Name & "')." & vbCrLf & _
               "Please check the column names in the source sheet against the configuration and the ExecutionLog.txt for details.", vbCritical
    End If
    Exit Function

ErrorHandler:
    Utils.LogMessage "[" & PROC_NAME & "] " & funcName & ": ERROR " & Err.Number & " - " & Err.Description, True
    GetColumnIndicesFromSource = False
End Function

Private Function SetupIQADatabase(ByRef iqaWorkbook As Workbook, ByVal cfg As Object) As Worksheet
    Dim iqaSource As Worksheet
    Dim iqaDatabasePath As String

    On Error GoTo ErrorHandler

        ' Retrieve the IQA Database path from the Config sheet
        iqaDatabasePath = ""
        If Not cfg Is Nothing And cfg.exists(Config.CONFIG_KEY_IQA_DB_PATH) Then
            iqaDatabasePath = cfg(Config.CONFIG_KEY_IQA_DB_PATH)
        End If

        If iqaDatabasePath = "" Then
            Utils.LogMessage "[" & PROC_NAME & "] IQA Database path key '" & Config.CONFIG_KEY_IQA_DB_PATH & "' not found or empty in configuration!", True
            Exit Function
        End If

        ' Open the IQA Database workbook
        Application.EnableEvents = False
        Set iqaWorkbook = Workbooks.Open(Filename:=iqaDatabasePath, ReadOnly:=True, IgnoreReadOnlyRecommended:=True, UpdateLinks:=False)
        Application.EnableEvents = True

        If iqaWorkbook Is Nothing Then
            Utils.LogMessage "[" & PROC_NAME & "] Failed to open IQA Database at: " & iqaDatabasePath, True
            Exit Function
        End If

        Set iqaSource = iqaWorkbook.Sheets(Config.IQA_SHEET_NAME) ' Use constant from Config.bas
        If iqaSource Is Nothing Then
            Utils.LogMessage "[" & PROC_NAME & "] Sheet '" & Config.IQA_SHEET_NAME & "' not found in the IQA workbook!", True
            MsgBox "Sheet '" & Config.IQA_SHEET_NAME & "' not found in the IQA workbook!", vbExclamation
            Exit Function
        End If

        ' remove active filter in the IQA Database sheet
        If iqaSource.AutoFilterMode Then iqaSource.AutoFilterMode = False

        Utils.LogMessage "[" & PROC_NAME & "] IQA Database setup successfully from: " & iqaDatabasePath, False
        
        ' After successfully setting wsSource, get the column indices
        If Not GetColumnIndicesFromSource(iqaSource) Then
            Set iqaSource = Nothing ' Indicate failure
            ' Error already logged by GetColumnIndicesFromSource
            Exit Function
        End If
        
        Set SetupIQADatabase = iqaSource

    Exit Function

ErrorHandler: ' Corrected label
        Utils.LogMessage "[" & PROC_NAME & "] Error setting up IQA Database: " & Err.Description, True
        ' iqaWorkbook will be closed in the main CleanUp block if it was opened
        Application.EnableEvents = True
        Set SetupIQADatabase = Nothing
End Function
