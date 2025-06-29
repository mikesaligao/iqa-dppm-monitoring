Attribute VB_Name = "Config"
Option Explicit

' --- Sheet Names ---
Public Const IQA_SHEET_NAME As String = "IQA Database"
Public Const ROUTING_SHEET_NAME As String = "Routing Database"

' --- Table Names ---
Public Const IQA_TABLE_NAME As String = "tblIQADatabase"
Public Const ROUTING_TABLE_NAME As String = "tblRoutingDatabase"

' --- IQA Database Table Columns ---
Public Const IQA_COL_CONTROL_NO As String = "Control No." ' Control No.
Public Const IQA_COL_SUPPLIER As String = "Supplier Name" ' Supplier Name
Public Const IQA_COL_PART_NUM As String = "Part Number" ' Part Number
Public Const IQA_COL_PART_DESC As String = "Part Description" ' Part Description
Public Const IQA_COL_PO_NUMBER As String = "PO Number" ' PO Number
Public Const IQA_COL_QUANTITY_IN As String = "Quantity In" ' Quantity In
Public Const IQA_COL_RESERVED_FIELD_1 As String = "Reserved Field 1" ' Reserved Field 1
Public Const IQA_COL_RESERVED_FIELD_2 As String = "Reserved Field 2" ' Reserved Field 2
Public Const IQA_COL_USER_NAME As String = "User Name" ' User Name
Public Const IQA_COL_POSTED_DATE As String = "Posted Date" ' Posted Date
Public Const IQA_COL_SHIP_DATE As String = "Shipment Date" ' Shipment Date
Public Const IQA_COL_LOT_BATCH_TOOL_NUMBER As String = "Lot/Batch/Tool Number" ' Lot/Batch/Tool Number
Public Const IQA_COL_DATE_ENDORSED As String = "Date Endorsed" ' Date Endorsed
Public Const IQA_COL_ENDORSED_BY As String = "Endorsed By" ' Endorsed By
Public Const IQA_COL_RECEIVED_BY As String = "Received By" ' Received By
Public Const IQA_COL_WORK_WEEK_PER_SD As String = "Work Week (per SD)" ' Work Week (per SD)
Public Const IQA_COL_FAI_CONTROL_NO As String = "FAI Control No." ' FAI Control No.
Public Const IQA_COL_REQUESTED_BY As String = "Requested By" ' Requested By
Public Const IQA_COL_REMARKS As String = "Remarks" ' Remarks
Public Const IQA_COL_PROD_LINE As String = "Product Line" ' Production Line
Public Const IQA_COL_INSP_STATUS As String = "Inspection Type (DTS/FAI/Normal)" ' Inspection Status (FAI/Normal/DTS)
Public Const IQA_COL_IMMEDIATE_RELEASE_REQUESTED As String = "Immediate Release Requested?" ' Immediate Release Requested?
Public Const IQA_COL_IQE_DISPOSITION As String = "IQE Disposition" ' IQE Disposition
Public Const IQA_COL_VISUAL_SAMPLE_SIZE As String = "Visual Sample Size" ' Visual Sample Size
Public Const IQA_COL_COMMIT_DATE As String = "Commit Date" ' Commit Date
Public Const IQA_COL_RECOMMIT_DATE As String = "Recommit Date" ' Recommit Date
Public Const IQA_COL_INSPECTION_START_DATE As String = "Inspection Start Date" ' Inspection Start Date
Public Const IQA_COL_INSPECTION_END_DATE_EXCL_FT As String = "Inspection End Date (Excl. FT)" ' Inspection End Date (Excl. FT)
Public Const IQA_COL_WORK_WEEK_COMPLETED As String = "Work Week Completed" ' Work Week Completed
Public Const IQA_COL_IQA_OCD_CRD_COUNT As String = "IQA OCD/CRD Count" ' IQA OCD/CRD Count
Public Const IQA_COL_VISUAL_INSPECTION_START As String = "Visual Inspection Start" ' Visual Inspection Start
Public Const IQA_COL_VISUAL_INSPECTION_END As String = "Visual Inspection End" ' Visual Inspection End
Public Const IQA_COL_INSPECTED_BY As String = "Inspected by" ' Inspected by
Public Const IQA_COL_VISUAL_INSPECTION_RESULT As String = "Visual Inspection Result" ' Visual Inspection Result
Public Const IQA_COL_VISUAL_REJECT_QUANTITY As String = "Visual Reject Quantity" ' Visual Reject Quantity
Public Const IQA_COL_DIMENSIONAL_INSPECTION_START As String = "Dimensional Inspection Start" ' Dimensional Inspection Start
Public Const IQA_COL_DIMENSIONAL_INSPECTION_END As String = "Dimensional Inspection End" ' Dimensional Inspection End
Public Const IQA_COL_DIMENSIONAL_RESULT As String = "Dimensional Result" ' Dimensional Result
Public Const IQA_COL_DIMENSIONAL_INSPECTOR As String = "Dimensional Inspector" ' Dimensional Inspector
Public Const IQA_COL_FUNCTIONAL_TEST_ENDORSED_DATE As String = "Functional Test Endorsed Date" ' Functional Test Endorsed Date
Public Const IQA_COL_RETURN_DATE_TO_IQA As String = "Return Date To IQA" ' Return Date To IQA
Public Const IQA_COL_FUNCTIONAL_TEST_ENDORSED_TO As String = "Functional Test Endorsed To" ' Functional Test Endorsed To
Public Const IQA_COL_FT_ENDORSED_BY As String = "FT Endorsed By" ' FT Endorsed By
Public Const IQA_COL_FUNCTIONAL_TEST_RESULT As String = "Functional Test Result" ' Functional Test Result
Public Const IQA_COL_INITIAL_DISPOSITION As String = "Initial Disposition" ' Initial Disposition
Public Const IQA_COL_FINAL_DISP As String = "Final Disposition" ' Final Disposition (Accept/Reject)
Public Const IQA_COL_IQA_STATUS As String = "IQA Status" ' IQA Status
Public Const IQA_COL_OVERALL_STATUS As String = "Overall Status" ' Overall Status
Public Const IQA_COL_FINAL_INSPECTOR As String = "Final Inspector" ' Final Inspector
Public Const IQA_COL_DATE_RELEASED As String = "Date Released" ' Date Released
Public Const IQA_COL_ENDORSED_TO As String = "Endorsed To" ' Endorsed To
Public Const IQA_COL_WORKDAYS_PENDING_AT_IQA As String = "Workdays Pending at IQA" ' Workdays Pending at IQA
Public Const IQA_COL_CYCLE_TIME As String = "Cycle Time" ' Cycle Time
Public Const IQA_COL_ADDITIONAL_REMARKS As String = "Additional Remarks" ' Additional Remarks
Public Const IQA_COL_TOTAL_REJECT_QUANTITY As String = "Total Reject Quantity" ' Total Reject Quantity
Public Const IQA_COL_DEFECT_DETECTION_DATE As String = "Defect Detection Date" ' Defect Detection Date
Public Const IQA_COL_DEFECT_TYPE As String = "Defect Type" ' Defect Type
Public Const IQA_COL_DEFECT_AREA_IQA_LINE As String = "Defect Area (IQA/Line)" ' Defect Area (IQA/Line)
Public Const IQA_COL_PROD_ISSUE_NOTES As String = "Production Issue Notes" ' Free-form text For production issues

Public Const STATUS_BLANK_MARKER As String = "---" ' Marker For blank status in IQA

' --- Security ---
Public Const CONST_SHEET_PROTECTION_PASSWORD As String = "IQA2025" ' Change "YourSecurePassword" To your actual password

' --- DPPM Generation ---
Public Const DPPM_OUTPUT_SHEET_NAME As String = "dppm-database"
Public Const DPPM_OUTPUT_TABLE_NAME As String = "tblDppmDatabase"
Public Const DPPM_COL_DATE As String = "Date"
Public Const DPPM_COL_SUPPLIER As String = "Supplier Name"
Public Const DPPM_COL_PART_NUM As String = "Part Number"
Public Const DPPM_COL_INSPECTED_BY As String = "Inspected By"
Public Const DPPM_COL_OVERALL_QTY As String = "Overall Quantity Received"
Public Const DPPM_COL_OVERALL_REJECT As String = "Overall Units Reject"
Public Const DPPM_COL_OVERALL_DPPM As String = "Overall DPPM"
Public Const DPPM_COL_INSPECTED_QTY As String = "Inspected Quantity Received"
Public Const DPPM_COL_INSPECTED_REJECT As String = "Inspected Units Reject"
Public Const DPPM_COL_INSPECTED_DPPM As String = "Inspected DPPM"

' --- Wafer List Table and Columns ---
Public Const WAFER_LIST_SHEET_NAME As String = "Wafer List"
Public Const WAFER_LIST_TABLE_NAME As String = "tblWaferList"
Public Const WAFER_LIST_COL_PART_NUM As String = "Part Number"
Public Const WAFER_LIST_COL_PART_DESC As String = "Part Description"
Public Const WAFER_LIST_COL_CHIPS_PER_WAFER As String = "Number of chips per wafer"

Public Const CONFIG_KEY_IQA_DB_PATH As String = "IQA Database Path"
Public Const CONFIG_KEY_IQA_SRC_SHIP_DATE_COLNAME As String = "IQA Source Shipment Date ColName"
Public Const CONFIG_KEY_IQA_SRC_INSP_DATE_COLNAME As String = "IQA Source Inspected Date ColName"
Public Const CONFIG_KEY_IQA_SRC_SUPPLIER_COLNAME As String = "IQA Source Supplier Name ColName"
Public Const CONFIG_KEY_IQA_SRC_PARTNUM_COLNAME As String = "IQA Source Part Number ColName"
Public Const CONFIG_KEY_IQA_SRC_INSP_BY_COLNAME As String = "IQA Source Inspected By ColName"
Public Const CONFIG_KEY_IQA_SRC_QTY_IN_COLNAME As String = "IQA Source Quantity In ColName"
Public Const CONFIG_KEY_IQA_SRC_REJ_QTY_COLNAME As String = "IQA Source Reject Quantity ColName"

' --- DPPM Summary ---
' Keys for Config Sheet (to retrieve sheet/table names for summaries)
Public Const CONFIG_KEY_DPPM_DAILY_SHEET_NAME As String = "DailySummary"
Public Const CONFIG_KEY_DPPM_DAILY_TABLE_NAME As String = "tblDailySummary"
Public Const CONFIG_KEY_DPPM_WEEKLY_SHEET_NAME As String = "WeeklySummary"
Public Const CONFIG_KEY_DPPM_WEEKLY_TABLE_NAME As String = "tblWeeklySummary"
Public Const CONFIG_KEY_DPPM_MONTHLY_SHEET_NAME As String = "MonthlySummary"
Public Const CONFIG_KEY_DPPM_MONTHLY_TABLE_NAME As String = "tblMonthlySummary"

' Column Names for Summary Tables
Public Const SUMMARY_COL_PERIOD As String = "Date" ' Or "Period"; holds Day, Week (YYYY-WW##), Month (YYYY-MMMM)
Public Const SUMMARY_COL_OVERALL_QTY As String = "Overall Qty Received"
Public Const SUMMARY_COL_OVERALL_REJECT As String = "Overall Units Reject"
Public Const SUMMARY_COL_OVERALL_DPPM_CALC As String = "Overall DPPM"
Public Const SUMMARY_COL_INSPECTED_QTY As String = "Inspected Qty Received"
Public Const SUMMARY_COL_INSPECTED_REJECT As String = "Inspected Units Reject"
Public Const SUMMARY_COL_INSPECTED_DPPM_CALC As String = "Inspected DPPM"

Public Const DEFAULT_TABLE_STYLE As String = "TableStyleMedium9"

' --- Procedure Name Constants (for Logging/Status) ---
Public Const PROC_GENERATE_SUMMARY As String = "GenerateDPPMSummary"
Public Const PROC_GENERATE_SUMMARY_BY_TYPE As String = "GenerateSummaryByType"
Public Const PROC_LOAD_SUMMARY_CONFIG As String = "LoadSummaryConfig"

' --- Procedure Name Constants (for Logging/Status) ---
Public Const PROC_GENERATE_TABLE As String = "GenerateDPPMTable"
Public Const PROC_GENERATE_TABLE_WRITE As String = "GenerateTable_WriteDPPMTable"
Public Const PROC_GENERATE_TABLE_FORMAT As String = "GenerateTable_FormatDPPMTable"
Public Const PROC_LOAD_IQA_DB As String = "GenerateTable_LoadIQADatabase"
