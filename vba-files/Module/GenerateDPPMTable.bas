Attribute VB_Name = "GenerateDPPMTable"

Option Explicit
' Declare a global variable To store configuration
Public GlobalConfig As Object

Public ShipmentDateCol As Long ' Column For Shipment Date
Public InspectedDateCol As Long ' Column For Inspected Date
Public SupplierNameCol As Long ' Column For Supplier Name
Public PartNumberCol As Long ' Column For Part Number
Public InspectedByCol As Long ' Column For Inspected By
Public QuantityInCol As Long ' Column For Quantity In
Public RejectQuantityCol As Long ' Column For Reject Quantity

Public Sub GenerateTable()
    Dim wsSource As Worksheet, wsTarget As Worksheet, wsWafer As Worksheet
    Dim lastRow As Long, targetRow As Long
    Dim key As Variant
    Dim shipmentDate As String, inspectedDate As String
    Dim supplierName As String, partNumber As String, inspectedBy As String
    Dim quantityIn As Double, rejectQuantity As Double
    Dim overallQuantity As Double, overallRejects As Double, overallDPPM As Double
    Dim inspectedQuantity As Double, inspectedRejects As Double, inspectedDPPM As Double
    Dim tempArray As Variant
    Dim chipsPerWaferCount As Double
    Dim dataDict As Object
    Dim i As Long
    Dim iqaWorkbook As Workbook
    Dim outputArr() As Variant

    On Error GoTo ErrorHandler

        ' Disable screen updating For performance
        Application.ScreenUpdating = False
        Debug.Print "Execution started: " & Now

        ' Load configuration from the Config sheet
        Debug.Print "Loading configuration from Config sheet."
        Set GlobalConfig = Nothing
        Call LoadConfigFromSheet

        ' Set the source sheet from the current workbook
        Set wsSource = SetupIQADatabase(iqaWorkbook)
        If wsSource Is Nothing Then Exit Sub
            Debug.Print "Source sheet 'IQA Database' found in ThisWorkbook."

            ' Check If the "dppm-database" sheet exists in ThisWorkbook
            On Error Resume Next
            Set wsTarget = ThisWorkbook.Sheets("dppm-database")
            On Error GoTo 0
                If wsTarget Is Nothing Then
                    Debug.Print "Target sheet 'dppm-database' does Not exist in ThisWorkbook. Creating a New sheet."
                    Set wsTarget = ThisWorkbook.Sheets.Add
                    wsTarget.Name = "dppm-database"
                Else
                    Debug.Print "Target sheet 'dppm-database' found in ThisWorkbook."
                End If

                ' Set the Wafer List sheet from the current workbook
                On Error Resume Next
                Set wsWafer = ThisWorkbook.Sheets("Wafer List")
                On Error GoTo 0
                    If wsWafer Is Nothing Then
                        MsgBox "Wafer List sheet Not found!", vbExclamation
                     Exit Sub
                    End If

                    ' Initialize dictionary For aggregation
                    Debug.Print "Initializing dictionary For data aggregation."
                    Set dataDict = CreateObject("Scripting.Dictionary")

                    ' Find the last row in the source sheet
                    lastRow = wsSource.Cells(wsSource.Rows.Count, "B").End(xlUp).Row
                    Debug.Print "Last row in source sheet: " & lastRow

                    ' Read the entire range into an array To avoid repeated cell access
                    Dim dataArr As Variant
                    ' Assuming the data starts from row 2 And goes To the last row in column BC
                    dataArr = wsSource.Range(wsSource.Cells(2, 1), wsSource.Cells(lastRow, RejectQuantityCol)).Value

                    ' Loop through each row in the source sheet
                    Debug.Print "Starting data extraction And aggregation."
                    For i = 1 To UBound(dataArr, 1) ' Start from 1 since dataArr is 1-based
                        shipmentDate = Format(dataArr(i, ShipmentDateCol), "yyyy-MM-dd")
                        inspectedDate = dataArr(i, InspectedDateCol)
                        supplierName = dataArr(i, SupplierNameCol)
                        partNumber = dataArr(i, PartNumberCol)
                        inspectedBy = dataArr(i, InspectedByCol)

                        ' Quantity in is in column F, check If it's numeric
                        quantityIn = 0
                        If IsNumeric(dataArr(i, QuantityInCol)) Then
                            quantityIn = dataArr(i, QuantityInCol)
                        End If

                        ' Reject quantity is in column BC, check If it's numeric
                        rejectQuantity = 0
                        If IsNumeric(dataArr(i, RejectQuantityCol)) Then
                            rejectQuantity = dataArr(i, RejectQuantityCol)
                        End If

                        ' Log the raw values being read
                        Debug.Print "Row " & i & ": ShipmentDate=" & shipmentDate & ", SupplierName=" & supplierName & ", PartNumber=" & partNumber & ", InspectedBy=" & inspectedBy & ", QuantityIn=" & quantityIn & ", RejectQuantity=" & rejectQuantity

                        ' Reset chipsPerWaferCount To avoid reusing previous values
                        chipsPerWaferCount = 0

                        ' Check For supplier-specific logic
                        If supplierName = "EXCELITAS CANADA INC." Then
                            Debug.Print "Row " & i & ": Supplier is EXCELITAS CANADA INC. Checking Wafer List For chips per wafer count."
                            On Error Resume Next
                            chipsPerWaferCount = Application.WorksheetFunction.VLookup(partNumber, wsWafer.Range("A:C"), 3, False)
                            On Error GoTo 0
                                If chipsPerWaferCount <= 0 Then
                                    Debug.Print "Row " & i & ": Chips per wafer count Not found Or invalid. Skipping row."
                                    GoTo NextRow
                                    End If
                                    Debug.Print "Row " & i & ": Chips per wafer count found: " & chipsPerWaferCount
                                    quantityIn = quantityIn * chipsPerWaferCount
                                    Debug.Print "Row " & i & ": Adjusted Quantity In: " & quantityIn
                                End If

                                ' Skip rows With missing key data
                                If shipmentDate = "" Or supplierName = "" Or partNumber = "" Then
                                    Debug.Print "Row " & i & ": Missing key data. Skipping row."
                                    GoTo NextRow
                                    End If

                                    ' Create a unique key For grouping
                                    key = shipmentDate & "|" & supplierName & "|" & partNumber

                                    ' Aggregate data in the dictionary
                                    If Not dataDict.exists(key) Then
                                        Debug.Print "Adding New key To dictionary: " & key
                                        dataDict.Add key, Array(shipmentDate, supplierName, partNumber, inspectedBy, 0, 0, 0, 0)
                                    End If

                                    ' Retrieve, update, And store the array back in the dictionary
                                    tempArray = dataDict(key)
                                    tempArray(4) = tempArray(4) + quantityIn ' Overall Quantity Received
                                    tempArray(5) = tempArray(5) + rejectQuantity ' Overall Units Reject
                                    dataDict(key) = tempArray

                                    Debug.Print "Row " & i & ": Overall Quantity Received updated: " & tempArray(4)
                                    Debug.Print "Row " & i & ": Overall Units Reject updated: " & tempArray(5)

                                    ' Check If inspectedDate is a valid date
                                    If IsDate(inspectedDate) Then
                                        inspectedDate = Format(inspectedDate, "yyyy-MM-dd")
                                        key = inspectedDate & "|" & supplierName & "|" & partNumber
                                        If Not dataDict.exists(key) Then
                                            Debug.Print "Adding New key To dictionary: " & key
                                            dataDict.Add key, Array(inspectedDate, supplierName, partNumber, inspectedBy, 0, 0, 0, 0)
                                        End If
                                        tempArray = dataDict(key)
                                        tempArray(6) = tempArray(6) + quantityIn ' Inspected Quantity Received
                                        tempArray(7) = tempArray(7) + rejectQuantity ' Inspected Units Reject
                                        dataDict(key) = tempArray

                                        Debug.Print "Row " & i & ": Inspected Quantity Received updated: " & tempArray(6)
                                        Debug.Print "Row " & i & ": Inspected Units Reject updated: " & tempArray(7)
                                    Else
                                        Debug.Print "Row " & i & ": Invalid Inspected Date. Skipping row."
                                    End If

NextRow:
                                Next i

                                ' Pre-size the output array: [1 To rowCount, 1 To 10]
                                ReDim outputArr(1 To dataDict.Count + 1, 1 To 10)

                                ' Write headers To the target sheet
                                Debug.Print "Writing headers To the target sheet."
                                wsTarget.Cells.Clear
                                outputArr(1, 1) = "Date"
                                outputArr(1, 2) = "Supplier Name"
                                outputArr(1, 3) = "Part Number"
                                outputArr(1, 4) = "Inspected By"
                                outputArr(1, 5) = "Overall Quantity Received"
                                outputArr(1, 6) = "Overall Units Reject"
                                outputArr(1, 7) = "Overall DPPM"
                                outputArr(1, 8) = "Inspected Quantity Received"
                                outputArr(1, 9) = "Inspected Units Reject"
                                outputArr(1, 10) = "Inspected DPPM"

                                ' Write aggregated data To the target sheet
                                Debug.Print "Writing aggregated data To the target sheet."
                                targetRow = 2
                                For Each key In dataDict.keys
                                    tempArray = dataDict(key)
                                    shipmentDate = tempArray(0)
                                    supplierName = tempArray(1)
                                    partNumber = tempArray(2)
                                    inspectedBy = tempArray(3)
                                    overallQuantity = tempArray(4)
                                    overallRejects = tempArray(5)
                                    inspectedQuantity = tempArray(6)
                                    inspectedRejects = tempArray(7)
                                    overallDPPM = 0
                                    If overallQuantity > 0 Then
                                        overallDPPM = (overallRejects / overallQuantity) * 1000000
                                    End If

                                    inspectedDPPM = 0
                                    If inspectedQuantity > 0 Then
                                        inspectedDPPM = (inspectedRejects / inspectedQuantity) * 1000000
                                    End If

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

                                    Debug.Print "Row " & targetRow & ": Data written For key " & key
                                    targetRow = targetRow + 1
                                Next key

                                ' Write all at once To the sheet, starting at row 1
                                wsTarget.Range("A1").Resize(dataDict.Count + 1, 10).Value = outputArr

                                ' Sort the data by Shipment Date
                                Debug.Print "Sorting data by Shipment Date."
                                wsTarget.Sort.SortFields.Clear
                                wsTarget.Sort.SortFields.Add key:=wsTarget.Columns(1), Order:=xlAscending
                                With wsTarget.Sort
                                    .SetRange wsTarget.UsedRange
                                    .Header = xlYes
                                    .MatchCase = False
                                    .Orientation = xlTopToBottom
                                    .SortMethod = xlPinYin
                                    .Apply
                                End With

                                ' Auto-fit columns
                                Debug.Print "Auto-fitting columns."
                                wsTarget.Columns("A:J").AutoFit
                                wsTarget.Columns("A:J").HorizontalAlignment = xlCenter
                                wsTarget.Columns("A:J").VerticalAlignment = xlCenter

                                ' Apply Borders
                                Debug.Print "Applying borders."
                                With wsTarget.Range("A1:J" & targetRow - 1).Borders
                                    .LineStyle = xlContinuous
                                    .ColorIndex = 0
                                    .TintAndShade = 0
                                    .Weight = xlThin
                                End With

                                Debug.Print "Execution completed: " & Now

                                ' Database info is updated, Update the Summary sheet
                                Call GenerateSummary
                                Debug.Print "Summary sheet updated."

                                MsgBox "DPPM table generated successfully!", vbInformation

Cleanup:
                                ' cleanup iqaWorkbook
                                If Not iqaWorkbook Is Nothing Then
                                    iqaWorkbook.Close SaveChanges:=False
                                    Set iqaWorkbook = Nothing

                                End If

                                ' Release objects To free memory
                                Set wsSource = Nothing
                                Set wsTarget = Nothing
                                Set wsWafer = Nothing
                                Set dataDict = Nothing

                                ' Re-enable screen updating
                                Application.ScreenUpdating = True
                             Exit Sub

ErrorHandler:
                                ' Display error message And Resume cleanup
                                MsgBox "An error occurred: " & Err.Description, vbCritical
                                Resume Cleanup
End Sub

Private Sub LoadConfigFromSheet()
    Dim wsConfig As Worksheet
    Dim lastRow As Long, i As Long
    Dim configDict As Object

    ' Set the Config sheet
    On Error Resume Next
    Set wsConfig = ThisWorkbook.Sheets("Config")
    On Error GoTo 0
        If wsConfig Is Nothing Then
            MsgBox "Config sheet Not found!", vbExclamation
         Exit Sub
        End If

        ' Initialize dictionary To store configuration
        Set configDict = CreateObject("Scripting.Dictionary")

        ' Find the last row in the Config sheet
        lastRow = wsConfig.Cells(wsConfig.Rows.Count, 1).End(xlUp).Row

        ' Loop through each row in the Config sheet
        For i = 2 To lastRow ' Assuming row 1 contains headers
            If wsConfig.Cells(i, 1).Value <> "" Then
                configDict(wsConfig.Cells(i, 1).Value) = wsConfig.Cells(i, 2).Value
            End If
        Next i

        ' Example: Accessing configuration values
        If configDict.exists("IQA Database Path") Then
            Debug.Print "IQA Database Path: " & configDict("IQA Database Path")
        End If

        If configDict.exists("Shipment Date Column") Then
            ShipmentDateCol = configDict("Shipment Date Column")
        End If

        If configDict.exists("Inspected Date Column") Then
            InspectedDateCol = configDict("Inspected Date Column")
        End If

        If configDict.exists("Supplier Name Column") Then
            SupplierNameCol = configDict("Supplier Name Column")
        End If

        If configDict.exists("Part Number Column") Then
            PartNumberCol = configDict("Part Number Column")
        End If

        If configDict.exists("Inspected By Column") Then
            InspectedByCol = configDict("Inspected By Column")
        End If

        If configDict.exists("Quantity In Column") Then
            QuantityInCol = configDict("Quantity In Column")
        End If

        If configDict.exists("Reject Quantity Column") Then
            RejectQuantityCol = configDict("Reject Quantity Column")
        End If

        ' Store the configuration dictionary in a global variable For reuse
        Set GlobalConfig = configDict

        Debug.Print "Configuration loaded successfully!"
End Sub

Private Function SetupIQADatabase(ByRef iqaWorkbook) As Worksheet
    Dim iqaSource As Worksheet
    Dim iqaDatabasePath As String

    On Error GoTo ErrorHandler

        ' Retrieve the IQA Database path from the Config sheet
        iqaDatabasePath = ""
        If Not GlobalConfig Is Nothing Then
            If GlobalConfig.exists("IQA Database Path") Then
                iqaDatabasePath = GlobalConfig("IQA Database Path")
            End If
        End If

        If iqaDatabasePath = "" Then
            MsgBox "IQA Database path Not found in Config sheet!", vbExclamation
         Exit Function
        End If

        ' Open the IQA Database workbook
        Application.EnableEvents = False
        Set iqaWorkbook = Workbooks.Open(Filename:=iqaDatabasePath, ReadOnly:=True, IgnoreReadOnlyRecommended:=True, UpdateLinks:=False)
        Application.EnableEvents = True

        If iqaWorkbook Is Nothing Then
            MsgBox "Failed To open IQA Database at: " & iqaDatabasePath, vbExclamation
         Exit Function
        End If

        Set iqaSource = iqaWorkbook.Sheets("IQA Database")
        If iqaSource Is Nothing Then
            MsgBox "IQA Database sheet Not found in the workbook!", vbExclamation
            iqaWorkbook.Close SaveChanges:=False
         Exit Function
        End If

        ' remove active filter in the IQA Database sheet
        iqaSource.AutoFilterMode = False

        Debug.Print "IQA Database setup successfully!"
        Set SetupIQADatabase = iqaSource

     Exit Function

ErrorHandler:
        MsgBox "An error occurred While setting up IQA Database: " & Err.Description, vbCritical
        If Not iqaWorkbook Is Nothing Then
            iqaWorkbook.Close SaveChanges:=False
        End If
        Application.EnableEvents = True
     Exit Function
End Function
