Attribute VB_Name = "Utils"
Option Explicit

' --- Logging System State ---
Private logFileInitialized As Boolean

' --- Status Bar Constants ---
Public Const STATUS_BAR_UPDATE_INTERVAL_SECONDS As Long = 1
Public Const STATUS_BAR_RECORD_UPDATE_INTERVAL As Long = 500 ' Records To process before an update
Public Const STATUS_BAR_MIN_ETA_SECONDS_THRESHOLD As Long = 5 ' Minimum elapsed time before showing ETA
Public Const STATUS_BAR_PROGRESS_BAR_LENGTH As Long = 20 ' Length of the text-based progress bar
Public Const STATUS_BAR_SPINNER_CHARS As String = "|/-\"

' --- Status Bar Global Variables ---
Public g_lngLastStatusBarUpdateTime As Double ' Timer value of the last status bar update
Private g_intSpinnerIndex As Integer         ' Current index For the spinner animation
Public g_blnStatusBarActive As Boolean      ' Tracks If status bar is actively managed by these routines (Public for external check/reset if absolutely necessary)
Private g_ConfigDict As Object             ' Globally accessible configuration dictionary

' --- Worksheet Handling ---

Function GetSheet(sheetName As String) As Worksheet
    On Error Resume Next
    Set GetSheet = ThisWorkbook.Sheets(sheetName)
    If Err.Number <> 0 Then
        LogMessage "Error getting sheet: '" & sheetName & "'. " & Err.Description, True
        Set GetSheet = Nothing
        Err.Clear
    End If
    On Error Goto 0
End Function

Function GetLastRow(ws As Worksheet, Optional columnRef As Variant = 1) As Long
    If ws Is Nothing Then
        GetLastRow = 1 ' Or 0 depending on how you handle header rows
     Exit Function
    End If
    On Error Resume Next ' In Case sheet is empty
    GetLastRow = ws.Cells(ws.Rows.Count, columnRef).End(xlUp).row
    If Err.Number <> 0 Then
        GetLastRow = 1 ' Default To 1 If error (e.g., empty sheet)
        Err.Clear
    End If
    On Error Goto 0
End Function

' --- Data Handling ---

Function CreateCompositeKey(partNumber As String, supplierName As String) As String
    ' Creates a consistent key, handling potential trims
    CreateCompositeKey = Trim(partNumber) & "-" & Trim(supplierName)
End Function

Function Nz(value As Variant, Optional defaultValue As Variant = "") As Variant
    ' Handles Null, Empty, Error, And Blank strings
    If isError(value) Then
        Nz = defaultValue
    Elseif IsEmpty(value) Or IsNull(value) Or Trim(CStr(value)) = "" Then
        Nz = defaultValue
    Else
        Nz = value
    End If
End Function

' --- Initialize Logging: Overwrite log file For each run ---
Public Sub InitLogFile()
    Dim logFilePath As String
    Dim FileNum As Integer
    Dim timeStamp As String

    On Error Goto ErrHandler

        logFilePath = ThisWorkbook.Path & "\ExecutionLog.txt"
        ' kill the logfile If its already opened
        If Dir(logFilePath) <> "" Then
            On Error Resume Next ' Ignore errors If file is open
            Kill logFilePath
            On Error Goto ErrHandler ' Restore error handling
            End If

            FileNum = FreeFile

            ' Overwrite the log file (Output mode)
            Open logFilePath For Output As #FileNum
            timeStamp = Format(Now, "yyyy-mm-dd hh:nn:ss")
            Print #FileNum, "=== New Session Started: " & timeStamp & " ==="
            Close #FileNum

            logFileInitialized = True
         Exit Sub

 ErrHandler:
            Debug.Print "[ERROR] Failed To initialize log file: " & Err.Description
            logFileInitialized = False
End Sub

' --- Logging ---
Public Sub LogMessage(message As String, Optional isError As Boolean = False)
    Dim logFilePath As String
    Dim FileNum As Integer
    Dim timeStamp As String

    On Error Resume Next ' Prevent errors during logging itself

    logFilePath = ThisWorkbook.Path & "\ExecutionLog.txt"
    FileNum = FreeFile
    timeStamp = Format(Now, "yyyy-mm-dd hh:nn:ss")

    ' Ensure log file is initialized For this session
    If Not logFileInitialized Then Call InitLogFile
        On Error Resume Next

        Open logFilePath For Append As #FileNum
        If Err.Number <> 0 Then Goto CleanUp ' Cannot open log file

            If isError Then
                Print #FileNum, "[ERROR] " & timeStamp & " - " & message
                Debug.Print "[ERROR] " & timeStamp & " - " & message
                MsgBox "Error: " & message, vbCritical, "Error"
            Else
                Print #FileNum, "[INFO]  " & timeStamp & " - " & message
                Debug.Print "[INFO]  " & timeStamp & " - " & message
            End If

 CleanUp:
            If FileNum > 0 Then Close #FileNum
                On Error Goto 0 ' Restore default error handling
End Sub

Public Sub DeleteExecutionLog()
    Dim logFilePath As String
    On Error Resume Next ' Prevent errors from stopping execution
    logFilePath = ThisWorkbook.Path & "\ExecutionLog.txt"
    If Dir(logFilePath) <> "" Then
        Kill logFilePath
        Debug.Print "ExecutionLog.txt has been deleted."
    Else
        Debug.Print "ExecutionLog.txt does Not exist."
    End If
    On Error Goto 0 ' Reset error handling
End Sub

' --- Worksheet Protection ---
Public Sub ProtectSheet(ws As Worksheet, Optional Byval password As String = "")
    If ws Is Nothing Then
        LogMessage "Cannot protect sheet: Worksheet object is Nothing.", True
     Exit Sub
    End If

    ' REMOVED: Since we are now using Excel Tables (ListObjects), we don't need To Set AutoFilter manually.
    '   If Not ws.AutoFilterMode Then
    '       ' AutoFilter Before protecting the sheet
    '       ws.Range(ROUTING_COL_PART_NUM & "1:" & ROUTING_COL_REMARKS & "1").AutoFilter _
    '      Field:=1, _
    '     VisibleDropDown:=True
    'End If

    On Error Goto ErrHandler
        If password <> "" Then
            ws.Protect password:=password, DrawingObjects:=True, Contents:=True, Scenarios:=True, _
            AllowSorting:=True, _
            AllowFiltering:=True, _
            AllowUsingPivotTables:=True
            LogMessage "Sheet '" & ws.Name & "' protected With a password."
        Else
            ws.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, _
            AllowSorting:=True, _
            AllowFiltering:=True, _
            AllowUsingPivotTables:=True
            LogMessage "Sheet '" & ws.Name & "' protected without a password."
        End If
     Exit Sub

 ErrHandler:
        LogMessage "Error protecting sheet '" & ws.Name & "'. " & Err.Description, True
        Err.Clear
End Sub

Public Sub UnprotectSheet(ws As Worksheet, Optional Byval password As String = "")
    If ws Is Nothing Then
        LogMessage "Cannot unprotect sheet: Worksheet object is Nothing.", True
     Exit Sub
    End If

    On Error Goto ErrHandler
        If password <> "" Then
            ws.Unprotect password:=password
            LogMessage "Sheet '" & ws.Name & "' unprotected With a password."
        Else
            ws.Unprotect
            LogMessage "Sheet '" & ws.Name & "' unprotected without a password."
        End If
     Exit Sub

 ErrHandler:
        LogMessage "Error unprotecting sheet '" & ws.Name & "'. " & Err.Description, True
        Err.Clear
End Sub

Public Function GetColumnIndexByName(tbl As ListObject, columnName As String) As Long
    On Error Resume Next
    GetColumnIndexByName = 0 ' Default To 0 If Not found Or error
    If tbl Is Nothing Then
        LogMessage "GetColumnIndexByName: Table object is Nothing. Cannot find column '" & columnName & "'.", True
     Exit Function
    End If
    GetColumnIndexByName = tbl.ListColumns(columnName).Index
    If Err.Number <> 0 Then
        LogMessage "GetColumnIndexByName: Column '" & columnName & "' Not found in table '" & tbl.Name & "'. Error: " & Err.Description, True
        Err.Clear
    End If
    On Error Goto 0
End Function

' --- Status Bar Helper Procedures ---

' Initializes the status bar For a procedure
Public Sub InitStatusBar(Byval procName As String)
    On Error Resume Next ' Ignore errors If status bar is unavailable
    If g_blnStatusBarActive And Application.StatusBar <> False Then Exit Sub ' Avoid nested control If already active

        Application.StatusBar = "Starting " & procName & "..."
        LogMessage "StatusBar: Starting " & procName & "..." ' Utils.LogMessage becomes LogMessage
        g_lngLastStatusBarUpdateTime = Timer
        g_intSpinnerIndex = 0
        g_blnStatusBarActive = True
        DoEvents ' Ensure initial message is displayed
End Sub

' Updates the status bar With a progress message, visual bar, percentage, And ETA
Public Sub UpdateStatusBarProgress(Byval operation As String, _
    Byval currentStep As Long, _
    Byval totalSteps As Long, _
    Optional Byval procStartTime As Double = 0, _
    Optional Byval subTask As String = "")
    On Error Resume Next
    If Not g_blnStatusBarActive Then Exit Sub

        Dim percentage As Double
        Dim progressBar As String
        Dim filledLength As Long
        Dim etaString As String
        Dim elapsedTime As Double
        Dim etaSeconds As Double
        Dim statusText As String

        If totalSteps > 0 Then
            If currentStep >= totalSteps Then
                percentage = 100
                filledLength = STATUS_BAR_PROGRESS_BAR_LENGTH
                progressBar = "[" & String(STATUS_BAR_PROGRESS_BAR_LENGTH, ChrW(&H25A0)) & "]" ' Full bar (solid squares)
            Elseif currentStep <= 0 Then
                percentage = 0
                filledLength = 0
                progressBar = "[" & String(STATUS_BAR_PROGRESS_BAR_LENGTH, ChrW(&H25A1)) & "]" ' Empty bar (empty squares)
            Else
                percentage = (CDbl(currentStep) / CDbl(totalSteps)) * 100
                filledLength = Application.WorksheetFunction.RoundDown(percentage / 100 * STATUS_BAR_PROGRESS_BAR_LENGTH, 0)
                If filledLength = 0 Then ' Progress started but Not enough For one block
                    progressBar = "[" & String(1, ChrW(&H25B6)) & String(STATUS_BAR_PROGRESS_BAR_LENGTH - 1, ChrW(&H25A1)) & "]" ' Triangle head, Then empty squares
                Elseif filledLength >= STATUS_BAR_PROGRESS_BAR_LENGTH Then
                    progressBar = "[" & String(STATUS_BAR_PROGRESS_BAR_LENGTH, ChrW(&H25A0)) & "]"
                Else
                    progressBar = "[" & String(filledLength, ChrW(&H25A0)) & String(1, ChrW(&H25B6)) & String(STATUS_BAR_PROGRESS_BAR_LENGTH - filledLength - 1, ChrW(&H25A1)) & "]"
                End If
            End If
            progressBar = progressBar & " " & Format(percentage, "0") & "%"
        Else ' Indeterminate progress
            progressBar = "[Processing...]"
            percentage = -1 ' Indicates indeterminate
        End If

        ' ETA Calculation
        If procStartTime > 0 And totalSteps > 0 And currentStep > 0 And percentage >= 0 And percentage < 100 Then
            elapsedTime = Timer - procStartTime
            If elapsedTime > STATUS_BAR_MIN_ETA_SECONDS_THRESHOLD Then
                Dim stepsPerSecond As Double
                stepsPerSecond = CDbl(currentStep) / elapsedTime
                If stepsPerSecond > 0 Then
                    etaSeconds = (CDbl(totalSteps) - CDbl(currentStep)) / stepsPerSecond
                    If etaSeconds >= 0 Then
                        If etaSeconds < 60 Then
                            etaString = " (ETA: " & Format(etaSeconds, "0") & "s)"
                        Elseif etaSeconds < 3600 Then
                            etaString = " (ETA: " & Format(etaSeconds / 60, "0") & "m " & Format(etaSeconds Mod 60, "0") & "s)"
                        Else
                            etaString = " (ETA: " & Format(etaSeconds / 3600, "0") & "h " & Format((etaSeconds Mod 3600) / 60, "0") & "m)"
                        End If
                    End If
                End If
            End If
        End If

        statusText = operation
        If subTask <> "" Then statusText = statusText & " - " & subTask
            statusText = statusText & ": " & progressBar
            If totalSteps > 0 Then statusText = statusText & " (" & currentStep & "/" & totalSteps & ")"
                statusText = statusText & etaString

                Application.StatusBar = statusText
                LogMessage "StatusBar: " & statusText ' Utils.LogMessage becomes LogMessage
                g_lngLastStatusBarUpdateTime = Timer
                DoEvents ' Allow UI To update And respond
End Sub

' Updates the status bar With a simple message, optionally With a spinner Or completion checkmark
Public Sub UpdateStatusBarMessage(Byval message As String, Optional Byval useSpinner As Boolean = False, Optional Byval stageComplete As Boolean = False)
    On Error Resume Next
    If Not g_blnStatusBarActive Then Exit Sub

        Dim fullMessage As String
        fullMessage = message

        If useSpinner Then
            g_intSpinnerIndex = (g_intSpinnerIndex Mod Len(STATUS_BAR_SPINNER_CHARS)) + 1
            fullMessage = message & " " & Mid$(STATUS_BAR_SPINNER_CHARS, g_intSpinnerIndex, 1)
        End If

        If stageComplete Then
            fullMessage = message & " ✓"
        End If

        Application.StatusBar = fullMessage
        LogMessage "StatusBar: " & fullMessage ' Utils.LogMessage becomes LogMessage
        g_lngLastStatusBarUpdateTime = Timer
        DoEvents
End Sub

' Resets the status bar after completion Or error
Public Sub ResetStatusBar(Optional Byval procName As String = "", Optional Byval errOccurred As Boolean = False, Optional Byval errDescription As String = "")
    On Error Resume Next
    If Not g_blnStatusBarActive And procName = "" And Not errOccurred Then
        If Application.StatusBar <> False Then Application.StatusBar = False
         Exit Sub
        End If

        If errOccurred Then
            Dim errMsg As String
            errMsg = "Error in " & procName & ": " & errDescription & ". Check logs."
            Application.StatusBar = errMsg
            LogMessage "StatusBar: " & errMsg ' Utils.LogMessage becomes LogMessage
        Elseif procName <> "" Then
            Dim successMsg As String
            successMsg = procName & " completed. Ready."
            Application.StatusBar = successMsg
            LogMessage "StatusBar: " & successMsg ' Utils.LogMessage becomes LogMessage
        Else
            Application.StatusBar = False ' Reset To default Excel status bar
            LogMessage "StatusBar: Reset." ' Utils.LogMessage becomes LogMessage
        End If
        g_blnStatusBarActive = False
        DoEvents
End Sub

' --- Global Configuration Loading ---
Public Function GetGlobalConfig() As Object
    ' Returns the global configuration dictionary, loading it if necessary.
    If g_ConfigDict Is Nothing Then
        Call LoadGlobalConfig
    End If
    Set GetGlobalConfig = g_ConfigDict
End Function

Private Sub LoadGlobalConfig()
    ' Loads configuration from the "Config" sheet into g_ConfigDict.
    Dim wsConfig As Worksheet
    Dim lastRow As Long, i As Long
    Dim procName As String: procName = "Utils.LoadGlobalConfig"

    On Error GoTo ErrorHandler

    Set wsConfig = ThisWorkbook.Sheets("Config") ' Assumes "Config" sheet exists in ThisWorkbook
    If wsConfig Is Nothing Then
        LogMessage "[" & procName & "] 'Config' sheet not found in ThisWorkbook!", True
        Set g_ConfigDict = CreateObject("Scripting.Dictionary") ' Return empty dict
        Exit Sub
    End If

    Set g_ConfigDict = CreateObject("Scripting.Dictionary")
    g_ConfigDict.CompareMode = vbTextCompare ' Case-insensitive keys

    lastRow = wsConfig.Cells(wsConfig.Rows.Count, 1).End(xlUp).Row
    For i = 2 To lastRow ' Assuming row 1 contains headers
        Dim keyStr As String, valStr As String
        keyStr = Trim(CStr(wsConfig.Cells(i, 1).Value))
        valStr = Trim(CStr(wsConfig.Cells(i, 2).Value))
        If keyStr <> "" Then g_ConfigDict(keyStr) = valStr
    Next i
    LogMessage "[" & procName & "] Configuration loaded: " & g_ConfigDict.Count & " items.", False
    Exit Sub
ErrorHandler:
    LogMessage "[" & procName & "] ERROR " & Err.Number & ": " & Err.Description, True
    Set g_ConfigDict = CreateObject("Scripting.Dictionary") ' Return empty dict on error
End Sub
