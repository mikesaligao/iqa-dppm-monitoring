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

' Helper function to get chips per wafer for a given part number from tblWaferList
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

Public Sub GenerateTable()
    Dim tblIQA As ListObject, wsTarget As Worksheet, wsWafer As Worksheet, tblTarget As ListObject, tblWafer As ListObject
    Dim lastRow As Long, targetRow As Long, i As Long
    Dim key As Variant, tempArray As Variant, dataArr As Variant, outputArr() As Variant
    Dim shipmentDate As String, inspectedDateVal As Variant, supplierName As String, partNumber As String, inspectedBy As String
    Dim quantityIn As Double, rejectQuantity As Double, chipsPerWaferCount As Double
    Dim overallQuantity As Double, overallRejects As Double, overallDPPM As Double
    Dim inspectedQuantity As Double, inspectedRejects As Double, inspectedDPPM As Double
    Dim dataDict As Object
    Dim iqaWorkbook As Workbook
    Dim procStartTime As Double
    Dim colIdxShipDate As Long, colIdxInspDate As Long, colIdxSupplier As Long, colIdxPartNum As Long, colIdxInspBy As Long, colIdxQtyIn As Long, colIdxRejQty As Long

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

        '2025-06-27 MS: Disable the following line as it is not needed anymore
        'If Not LoadTableModuleConfig(m_Config) Then GoTo CleanUp ' Error logged in LoadTableModuleConfig
        'Utils.UpdateStatusBarMessage "Configuration loaded.", stageComplete:=True

        ' Set the source table from the IQA Database
        Utils.UpdateStatusBarMessage "Setting up IQA Database...", True
        Set tblIQA = SetupIQADatabase(iqaWorkbook, m_Config)
        If tblIQA Is Nothing Then GoTo CleanUp
        Utils.UpdateStatusBarMessage "IQA Database setup complete.", stageComplete:=True

        ' Get column indices from table using Config.bas constants
        colIdxShipDate = tblIQA.ListColumns(Config.IQA_COL_SHIP_DATE).Index
        colIdxInspDate = tblIQA.ListColumns(Config.IQA_COL_INSPECTED_BY).Index 
        colIdxSupplier = tblIQA.ListColumns(Config.IQA_COL_SUPPLIER).Index
        colIdxPartNum = tblIQA.ListColumns(Config.IQA_COL_PART_NUM).Index
        colIdxInspBy = tblIQA.ListColumns(Config.IQA_COL_INSPECTED_BY).Index
        colIdxQtyIn = tblIQA.ListColumns(Config.IQA_COL_QUANTITY_IN).Index
        colIdxRejQty = tblIQA.ListColumns(Config.IQA_COL_TOTAL_REJECT_QUANTITY).Index

        ' Read all data rows from the table
        dataArr = tblIQA.DataBodyRange.Value
        lastRow = UBound(dataArr, 1)
        Utils.LogMessage "[" & PROC_NAME & "] Last row in source IQA table: " & lastRow

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

        ' Loop through each row in the source data array
        Utils.LogMessage "[" & PROC_NAME & "] Starting data extraction and aggregation."
        procStartTime = Timer
        For i = 1 To lastRow
            If i Mod Utils.STATUS_BAR_RECORD_UPDATE_INTERVAL = 0 Or Timer - Utils.g_lngLastStatusBarUpdateTime > Utils.STATUS_BAR_UPDATE_INTERVAL_SECONDS Then
                Utils.UpdateStatusBarProgress "Aggregating DPPM Data", i, lastRow, procStartTime
            End If

            shipmentDate = Format(dataArr(i, colIdxShipDate), "yyyy-MM-dd")
            inspectedDateVal = dataArr(i, colIdxInspDate)
            supplierName = Trim(CStr(dataArr(i, colIdxSupplier)))
            partNumber = Trim(CStr(dataArr(i, colIdxPartNum)))
            inspectedBy = Trim(CStr(dataArr(i, colIdxInspBy)))

            quantityIn = 0
            If IsNumeric(dataArr(i, colIdxQtyIn)) Then quantityIn = CDbl(dataArr(i, colIdxQtyIn))

            rejectQuantity = 0
            If IsNumeric(dataArr(i, colIdxRejQty)) Then rejectQuantity = CDbl(dataArr(i, colIdxRejQty))

            'Utils.LogMessage "[" & PROC_NAME & "] Row " & i + 1 & ": ShipDate=" & shipmentDate & ", Supp=" & supplierName & ", PN=" & partNumber & ", QtyIn=" & quantityIn & ", RejQty=" & rejectQuantity, False

            chipsPerWaferCount = 0
            If supplierName = "EXCELITAS CANADA INC." Then
                chipsPerWaferCount = GetChipsPerWafer(tblWafer, partNumber)
                If chipsPerWaferCount > 0 Then
                    quantityIn = quantityIn * chipsPerWaferCount
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
        If lastRow > 0 Then Utils.UpdateStatusBarProgress "Aggregating DPPM Data", lastRow, lastRow, procStartTime ' Final update
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
Private Function GetColumnIndicesFromSource(tblSource As ListObject) As Boolean
    ' Determines column indices based on names stored in module-level variables using Excel Table (ListObject).
    ' This should be called AFTER tblSource is set.
    Dim funcName As String: funcName = "GetColumnIndicesFromSource"
    Dim allIndicesFound As Boolean
    Dim col As ListColumn

    On Error GoTo ErrorHandler
    GetColumnIndicesFromSource = False ' Default to failure
    allIndicesFound = True

    If tblSource Is Nothing Then
        Utils.LogMessage "[" & PROC_NAME & "] " & funcName & ": Source table is not set.", True
        Exit Function
    End If

    ' Helper to get column index by name from ListObject
    Dim functionColIdx As Long
    functionColIdx = 0
    
    ' Use ListColumns to get indices
    On Error Resume Next
    m_SourceShipmentDateColIdx = tblSource.ListColumns(m_SourceShipmentDateColName).Index
    If m_SourceShipmentDateColIdx = 0 Then allIndicesFound = False
    m_InspectedDateColIdx = tblSource.ListColumns(m_InspectedDateColName).Index
    If m_InspectedDateColIdx = 0 Then allIndicesFound = False
    m_SupplierNameColIdx = tblSource.ListColumns(m_SupplierNameColName).Index
    If m_SupplierNameColIdx = 0 Then allIndicesFound = False
    m_PartNumberColIdx = tblSource.ListColumns(m_PartNumberColName).Index
    If m_PartNumberColIdx = 0 Then allIndicesFound = False
    m_InspectedByColIdx = tblSource.ListColumns(m_InspectedByColName).Index
    If m_InspectedByColIdx = 0 Then allIndicesFound = False
    m_QuantityInColIdx = tblSource.ListColumns(m_QuantityInColName).Index
    If m_QuantityInColIdx = 0 Then allIndicesFound = False
    m_RejectQuantityColIdx = tblSource.ListColumns(m_RejectQuantityColName).Index
    If m_RejectQuantityColIdx = 0 Then allIndicesFound = False
    On Error GoTo 0

    If allIndicesFound Then
        Utils.LogMessage "[" & PROC_NAME & "] " & funcName & ": All source column indices determined successfully (using table).", False
        GetColumnIndicesFromSource = True
    Else
        MsgBox "One or more required columns were not found in the IQA Database table ('" & tblSource.Name & "')." & vbCrLf & _
               "Please check the column names in the source table against the configuration and the ExecutionLog.txt for details.", vbCritical
    End If
    Exit Function

ErrorHandler:
    Utils.LogMessage "[" & PROC_NAME & "] " & funcName & ": ERROR " & Err.Number & " - " & Err.Description, True
    GetColumnIndicesFromSource = False
End Function

Private Function SetupIQADatabase(ByRef iqaWorkbook As Workbook, ByVal cfg As Object) As ListObject
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

        ' Add a logic that will disable all macros when opening the IQA Database
        ' This is to ensure that the IQA Database is opened in a safe mode
        Set iqaWorkbook = Workbooks.Open(Filename:=iqaDatabasePath, ReadOnly:=True, IgnoreReadOnlyRecommended:=True, UpdateLinks:=False)
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

        Dim tblIQA As ListObject
        Set tblIQA = Nothing ' Reset the table variable to ensure it's not stale
        Set tblIQA = iqaSource.ListObjects(Config.IQA_TABLE_NAME) ' Use constant from Config.bas
        If tblIQATable Is Nothing Then
            Utils.LogMessage "[" & PROC_NAME & "] Table '" & Config.IQA_TABLE_NAME & "' not found in the IQA sheet '" & Config.IQA_SHEET_NAME & "'.", True
            MsgBox "Table '" & Config.IQA_TABLE_NAME & "' not found in the IQA sheet '" & Config.IQA_SHEET_NAME & "'.", vbExclamation
            Exit Function
        End If

        ' remove active filter in the IQA Database table if it exists
        If tblIQA.AutoFilter.FilterMode Then
            tblIQA.AutoFilter.ShowAllData
        End If
        Utils.LogMessage "[" & PROC_NAME & "] IQA Database setup successfully from: " & iqaDatabasePath, False
        
        ' After successfully setting wsSource, get the column indices
        If Not GetColumnIndicesFromSource(tblIQA) Then
            Set tblIQA = Nothing ' Indicate failure
            ' Error already logged by GetColumnIndicesFromSource
            Exit Function
        End If
        
        Set SetupIQADatabase = tblIQA

    Exit Function

ErrorHandler: ' Corrected label
        Utils.LogMessage "[" & PROC_NAME & "] Error setting up IQA Database: " & Err.Description, True
        ' iqaWorkbook will be closed in the main CleanUp block if it was opened
        Application.EnableEvents = True
        Set SetupIQADatabase = Nothing
End Function
