Attribute VB_Name = "mod_CoreBridgeAnalyticsHarness"
Option Explicit

'===============================================================================
' MODULE : mod_CoreBridgeAnalyticsHarness
' DOMAINE / DOMAIN : Validation Harnesses
'
' FR
' Harnais de preuve du contrat Core Bridge Analytics sur des copies de test.
' N'appartient a aucun workflow produit et ne doit pas etre appele en usage normal.
'
' EN
' Proof harness for the Core Bridge Analytics contract on test copies.
' Is not production workflow code and must not run during normal use.
'
' CONTRATS / CONTRACTS : CoreBridgeAnalyticsHarness_Smoke
' CALLBACKS EXTERNES / EXTERNAL CALLBACKS : Aucun / None
'===============================================================================


Private gCoreBridgeAnalyticsTracePath As String

'------------------------------------------------------------------------------
' FR: Execute un smoke contractuel des warnings CoreBridge Analytics.
' EN: Runs a contractual smoke for CoreBridge Analytics warnings.
'------------------------------------------------------------------------------
Public Function CoreBridgeAnalyticsHarness_Smoke( _
    ByVal capturePath As String, _
    ByVal matrixPath As String) As String

    Dim wsCalc As Worksheet
    Dim wsWBS As Worksheet
    Dim tblCalc As ListObject
    Dim tblWBS As ListObject
    Dim mapCalc As Object
    Dim mapWBS As Object
    Dim messages As Collection
    Dim deadlineRow As Long
    Dim parentRow As Long
    Dim childRow As Long
    Dim taskTypeRow As Long
    Dim deadlineFloatValue As Variant

    On Error GoTo Fail

    gCoreBridgeAnalyticsTracePath = matrixPath & ".trace.txt"
    CoreBridgeAnalyticsHarness_DeleteFile gCoreBridgeAnalyticsTracePath
    CoreBridgeAnalyticsHarness_DeleteFile matrixPath
    CoreBridgeAnalyticsHarness_DeleteFile capturePath
    CoreBridgeAnalyticsHarness_Trace "01 enter"

    EnsurePlanningEventHistoryInfrastructure
    PlanningConsolePolicy_DisableNonInteractive
    EventHistory_SetShowInfo True
    EventHistory_SetLanguage "EN"
    ClearPlanningWarningAcknowledgements
    ClearPlanningEventHistory
    CoreBridgeAnalyticsHarness_Trace "02 reset ok"

    Set wsCalc = ThisWorkbook.Worksheets("CALC")
    Set wsWBS = ThisWorkbook.Worksheets("WBS")
    Set tblCalc = wsCalc.ListObjects("tbl_CALC")
    Set tblWBS = wsWBS.ListObjects("tbl_WBS")
    Set mapCalc = CanonicalIdentity_BuildColumnMap(tblCalc)
    Set mapWBS = CanonicalIdentity_BuildColumnMap(tblWBS)

    CoreBridgeAnalyticsHarness_Assert Not tblCalc.DataBodyRange Is Nothing, "tbl_CALC has rows"
    CoreBridgeAnalyticsHarness_Assert Not tblWBS.DataBodyRange Is Nothing, "tbl_WBS has rows"
    CoreBridgeAnalyticsHarness_Assert tblCalc.ListRows.Count >= 4, "tbl_CALC has at least four rows"
    CoreBridgeAnalyticsHarness_Assert tblWBS.ListRows.Count >= 4, "tbl_WBS has at least four rows"

    Set messages = New Collection

    deadlineRow = 1
    parentRow = 2
    childRow = 3
    taskTypeRow = 4

    CoreBridgeAnalyticsHarness_ResetSyntheticInputs tblWBS, mapWBS, tblCalc, mapCalc

    CoreBridgeAnalyticsHarness_PrepareDeadlineRow tblCalc, mapCalc, deadlineRow
    CalcBridge_ComputeDeadlineAnalytics tblCalc, mapCalc, messages
    deadlineFloatValue = tblCalc.DataBodyRange.Cells(deadlineRow, mapCalc("Deadline Float")).Value
    CoreBridgeAnalyticsHarness_Assert CDbl(deadlineFloatValue) = -5#, "Deadline Float write preserved"
    CoreBridgeAnalyticsHarness_Trace "03 deadline ok"

    CoreBridgeAnalyticsHarness_PrepareParentRows tblCalc, mapCalc, parentRow, childRow
    CoreBridgeAnalyticsHarness_Assert CoreBridgeAnalyticsHarness_CalcRowIsSummary(tblCalc, mapCalc, "CBH-PARENT"), "CBH-PARENT is summary in CALC"
    CalcBridge_ShowParentDateWarnings tblCalc, mapCalc, messages
    CoreBridgeAnalyticsHarness_Trace "04 parent warning ok"

    CoreBridgeAnalyticsHarness_PrepareTaskTypeRows tblWBS, mapWBS, tblCalc, mapCalc, taskTypeRow
    CalcBridge_AppendTaskTypeWarnings messages, tblWBS, mapWBS, tblCalc, mapCalc
    CoreBridgeAnalyticsHarness_Trace "05 task type warning ok"

    CalcBridge_AddAnalyticsTopologyWarning messages
    CoreBridgeAnalyticsHarness_Trace "06 analytics warning ok"

    PlanningConsolePolicy_EnableNonInteractive capturePath, "CoreBridgeAnalyticsHarness"
    CalcBridge_ShowPlanningConsole messages
    PlanningConsolePolicy_DisableNonInteractive
    CoreBridgeAnalyticsHarness_Trace "07 console capture ok"

    CoreBridgeAnalyticsHarness_Assert CoreBridgeAnalyticsHarness_CaptureHasText(capturePath, "CoreBridgeAnalyticsHarness"), "console captured analytics warnings"
    CoreBridgeAnalyticsHarness_Assert CoreBridgeAnalyticsHarness_EventTypeExists("DEADLINE_EXCEEDED"), "DEADLINE_EXCEEDED event logged"
    CoreBridgeAnalyticsHarness_Assert CoreBridgeAnalyticsHarness_EventTypeExists("PARENT_DATES_IGNORED"), "PARENT_DATES_IGNORED event logged"
    CoreBridgeAnalyticsHarness_Assert CoreBridgeAnalyticsHarness_CaptureHasText(capturePath, "Deadline exceeded"), "Deadline warning projected"
    CoreBridgeAnalyticsHarness_Assert CoreBridgeAnalyticsHarness_CaptureHasText(capturePath, "Dates entered on summary task"), "Parent warning projected"
    CoreBridgeAnalyticsHarness_Assert CoreBridgeAnalyticsHarness_CaptureHasText(capturePath, "% Progress entered on LOE"), "Task type warning projected"
    CoreBridgeAnalyticsHarness_Assert CoreBridgeAnalyticsHarness_CaptureHasText(capturePath, "Analytics not calculated: incomplete topological order"), "Analytics warning projected"

    CoreBridgeAnalyticsHarness_WriteMatrix matrixPath, capturePath
    CoreBridgeAnalyticsHarness_Trace "08 matrix written"

    CoreBridgeAnalyticsHarness_Smoke = "PASS"
    Exit Function

Fail:
    On Error Resume Next
    CoreBridgeAnalyticsHarness_Trace "FAIL " & Err.Number & " " & Err.Description
    PlanningConsolePolicy_DisableNonInteractive
    CoreBridgeAnalyticsHarness_Smoke = "FAIL: " & Err.Description

End Function

'------------------------------------------------------------------------------
' FR: Verifie le contrat Analytics Harness Calc Row Is Summary et signale toute divergence au harnais.
' EN: Verifies the Analytics Harness Calc Row Is Summary contract and reports any divergence to the harness.
'------------------------------------------------------------------------------

Private Function CoreBridgeAnalyticsHarness_CalcRowIsSummary( _
    ByVal tblCalc As ListObject, _
    ByVal mapCalc As Object, _
    ByVal idValue As String) As Boolean

    Dim dataArr As Variant
    Dim rowById As Object
    Dim parentIds As Object

    dataArr = tblCalc.DataBodyRange.Value
    Set rowById = Core_BuildRowById(dataArr, mapCalc)
    Set parentIds = Core_BuildParentIds(dataArr, mapCalc, rowById)

    CoreBridgeAnalyticsHarness_CalcRowIsSummary = parentIds.Exists(idValue)

End Function

'------------------------------------------------------------------------------
' FR: Verifie le contrat Analytics Harness Reset Synthetic Inputs et signale toute divergence au harnais.
' EN: Verifies the Analytics Harness Reset Synthetic Inputs contract and reports any divergence to the harness.
'------------------------------------------------------------------------------

Private Sub CoreBridgeAnalyticsHarness_ResetSyntheticInputs( _
    ByVal tblWBS As ListObject, _
    ByVal mapWBS As Object, _
    ByVal tblCalc As ListObject, _
    ByVal mapCalc As Object)

    Dim r As Long

    For r = 1 To tblCalc.ListRows.Count
        CoreBridgeAnalyticsHarness_ClearIfExists tblCalc, mapCalc, r, "Deadline"
        CoreBridgeAnalyticsHarness_ClearIfExists tblCalc, mapCalc, r, "Deadline Float"
        CoreBridgeAnalyticsHarness_ClearIfExists tblCalc, mapCalc, r, "Actual Start"
        CoreBridgeAnalyticsHarness_ClearIfExists tblCalc, mapCalc, r, "Actual Finish"
        CoreBridgeAnalyticsHarness_ClearIfExists tblCalc, mapCalc, r, "Forecast Start"
        CoreBridgeAnalyticsHarness_ClearIfExists tblCalc, mapCalc, r, "Forecast Finish"
        CoreBridgeAnalyticsHarness_ClearIfExists tblCalc, mapCalc, r, "Baseline Start"
        CoreBridgeAnalyticsHarness_ClearIfExists tblCalc, mapCalc, r, "Baseline Duration"
        CoreBridgeAnalyticsHarness_ClearIfExists tblCalc, mapCalc, r, "Baseline Finish"
        CoreBridgeAnalyticsHarness_ClearIfExists tblCalc, mapCalc, r, "IsSummary"
    Next r

    For r = 1 To tblWBS.ListRows.Count
        CoreBridgeAnalyticsHarness_ClearIfExists tblWBS, mapWBS, r, "% Progress"
        CoreBridgeAnalyticsHarness_ClearIfExists tblWBS, mapWBS, r, "Baseline Start"
        CoreBridgeAnalyticsHarness_ClearIfExists tblWBS, mapWBS, r, "Baseline Duration"
        CoreBridgeAnalyticsHarness_ClearIfExists tblWBS, mapWBS, r, "Baseline Finish"
    Next r

End Sub

'------------------------------------------------------------------------------
' FR: Verifie le contrat Analytics Harness Clear If Exists et signale toute divergence au harnais.
' EN: Verifies the Analytics Harness Clear If Exists contract and reports any divergence to the harness.
' FR - Effet de bord : efface uniquement les donnees ou objets cibles du contrat.
' EN - Side effect: clears only data or objects targeted by the contract.
'------------------------------------------------------------------------------

Private Sub CoreBridgeAnalyticsHarness_ClearIfExists( _
    ByVal tbl As ListObject, _
    ByVal mapObj As Object, _
    ByVal rowIdx As Long, _
    ByVal columnName As String)

    If mapObj Is Nothing Then Exit Sub
    If Not mapObj.Exists(columnName) Then Exit Sub
    tbl.DataBodyRange.Cells(rowIdx, mapObj(columnName)).ClearContents

End Sub

'------------------------------------------------------------------------------
' FR: Verifie le contrat Analytics Harness Prepare Deadline Row et signale toute divergence au harnais.
' EN: Verifies the Analytics Harness Prepare Deadline Row contract and reports any divergence to the harness.
' FR - Effet de bord : ecrit dans une table Excel detenue par le workflow.
' FR - Effet de bord : efface uniquement les donnees ou objets cibles du contrat.
' EN - Side effect: writes to an Excel table owned by the workflow.
' EN - Side effect: clears only data or objects targeted by the contract.
'------------------------------------------------------------------------------

Private Sub CoreBridgeAnalyticsHarness_PrepareDeadlineRow( _
    ByVal tblCalc As ListObject, _
    ByVal mapCalc As Object, _
    ByVal rowIdx As Long)

    CoreBridgeAnalyticsHarness_RequireColumn mapCalc, "ID"
    CoreBridgeAnalyticsHarness_RequireColumn mapCalc, "WBS"
    CoreBridgeAnalyticsHarness_RequireColumn mapCalc, "Task Name"
    CoreBridgeAnalyticsHarness_RequireColumn mapCalc, "Deadline"
    CoreBridgeAnalyticsHarness_RequireColumn mapCalc, "Deadline Float"
    CoreBridgeAnalyticsHarness_RequireColumn mapCalc, "Calculated Finish"

    tblCalc.DataBodyRange.Cells(rowIdx, mapCalc("ID")).Value = "CBH-DEADLINE"
    tblCalc.DataBodyRange.Cells(rowIdx, mapCalc("WBS")).Value = "31.12.1"
    tblCalc.DataBodyRange.Cells(rowIdx, mapCalc("Task Name")).Value = "Harness deadline task"
    tblCalc.DataBodyRange.Cells(rowIdx, mapCalc("Calculated Finish")).Value = DateSerial(2026, 1, 20)
    tblCalc.DataBodyRange.Cells(rowIdx, mapCalc("Deadline")).Value = DateSerial(2026, 1, 15)
    tblCalc.DataBodyRange.Cells(rowIdx, mapCalc("Deadline Float")).ClearContents

End Sub

'------------------------------------------------------------------------------
' FR: Verifie le contrat Analytics Harness Prepare Parent Rows et signale toute divergence au harnais.
' EN: Verifies the Analytics Harness Prepare Parent Rows contract and reports any divergence to the harness.
' FR - Effet de bord : ecrit dans une table Excel detenue par le workflow.
' EN - Side effect: writes to an Excel table owned by the workflow.
'------------------------------------------------------------------------------

Private Sub CoreBridgeAnalyticsHarness_PrepareParentRows( _
    ByVal tblCalc As ListObject, _
    ByVal mapCalc As Object, _
    ByVal parentRow As Long, _
    ByVal childRow As Long)

    CoreBridgeAnalyticsHarness_RequireColumn mapCalc, "ID"
    CoreBridgeAnalyticsHarness_RequireColumn mapCalc, "ParentID"
    CoreBridgeAnalyticsHarness_RequireColumn mapCalc, "WBS"
    CoreBridgeAnalyticsHarness_RequireColumn mapCalc, "Task Name"
    CoreBridgeAnalyticsHarness_RequireColumn mapCalc, "Actual Start"
    CoreBridgeAnalyticsHarness_RequireColumn mapCalc, "Actual Finish"
    CoreBridgeAnalyticsHarness_RequireColumn mapCalc, "Forecast Start"
    CoreBridgeAnalyticsHarness_RequireColumn mapCalc, "Forecast Finish"
    CoreBridgeAnalyticsHarness_RequireColumn mapCalc, "Baseline Start"
    CoreBridgeAnalyticsHarness_RequireColumn mapCalc, "Baseline Duration"
    CoreBridgeAnalyticsHarness_RequireColumn mapCalc, "Baseline Finish"
    CoreBridgeAnalyticsHarness_RequireColumn mapCalc, "IsSummary"

    tblCalc.DataBodyRange.Cells(parentRow, mapCalc("ID")).Value = "CBH-PARENT"
    tblCalc.DataBodyRange.Cells(parentRow, mapCalc("ParentID")).Value = vbNullString
    tblCalc.DataBodyRange.Cells(parentRow, mapCalc("WBS")).Value = "31.12.2"
    tblCalc.DataBodyRange.Cells(parentRow, mapCalc("Task Name")).Value = "Harness parent task"
    tblCalc.DataBodyRange.Cells(parentRow, mapCalc("IsSummary")).Value = True
    tblCalc.DataBodyRange.Cells(parentRow, mapCalc("Forecast Start")).Value = DateSerial(2026, 2, 1)

    tblCalc.DataBodyRange.Cells(childRow, mapCalc("ID")).Value = "CBH-CHILD"
    tblCalc.DataBodyRange.Cells(childRow, mapCalc("ParentID")).Value = "CBH-PARENT"
    tblCalc.DataBodyRange.Cells(childRow, mapCalc("WBS")).Value = "31.12.2.1"
    tblCalc.DataBodyRange.Cells(childRow, mapCalc("Task Name")).Value = "Harness child task"
    tblCalc.DataBodyRange.Cells(childRow, mapCalc("IsSummary")).Value = False

End Sub

'------------------------------------------------------------------------------
' FR: Verifie le contrat Analytics Harness Prepare Task Type Rows et signale toute divergence au harnais.
' EN: Verifies the Analytics Harness Prepare Task Type Rows contract and reports any divergence to the harness.
' FR - Effet de bord : ecrit dans une table Excel detenue par le workflow.
' EN - Side effect: writes to an Excel table owned by the workflow.
'------------------------------------------------------------------------------

Private Sub CoreBridgeAnalyticsHarness_PrepareTaskTypeRows( _
    ByVal tblWBS As ListObject, _
    ByVal mapWBS As Object, _
    ByVal tblCalc As ListObject, _
    ByVal mapCalc As Object, _
    ByVal rowIdx As Long)

    CoreBridgeAnalyticsHarness_RequireColumn mapWBS, "ID"
    CoreBridgeAnalyticsHarness_RequireColumn mapWBS, "WBS"
    CoreBridgeAnalyticsHarness_RequireColumn mapWBS, "Task Type"
    CoreBridgeAnalyticsHarness_RequireColumn mapWBS, "% Progress"
    CoreBridgeAnalyticsHarness_RequireColumn mapWBS, "Baseline Start"
    CoreBridgeAnalyticsHarness_RequireColumn mapCalc, "ID"
    CoreBridgeAnalyticsHarness_RequireColumn mapCalc, "Task Type"

    tblWBS.DataBodyRange.Cells(rowIdx, mapWBS("ID")).Value = "CBH-LOE"
    tblWBS.DataBodyRange.Cells(rowIdx, mapWBS("WBS")).Value = "31.12.3"
    tblWBS.DataBodyRange.Cells(rowIdx, mapWBS("Task Type")).Value = "Level of Effort"
    tblWBS.DataBodyRange.Cells(rowIdx, mapWBS("% Progress")).Value = 0.5
    tblWBS.DataBodyRange.Cells(rowIdx, mapWBS("Baseline Start")).Value = DateSerial(2026, 3, 1)

    tblCalc.DataBodyRange.Cells(rowIdx, mapCalc("ID")).Value = "CBH-LOE"
    tblCalc.DataBodyRange.Cells(rowIdx, mapCalc("Task Type")).Value = "Level of Effort"

End Sub

'------------------------------------------------------------------------------
' FR: Verifie le contrat Analytics Harness Write Matrix et signale toute divergence au harnais.
' EN: Verifies the Analytics Harness Write Matrix contract and reports any divergence to the harness.
'------------------------------------------------------------------------------

Private Sub CoreBridgeAnalyticsHarness_WriteMatrix( _
    ByVal matrixPath As String, _
    ByVal capturePath As String)

    Dim fileNo As Integer

    fileNo = FreeFile
    Open matrixPath For Output As #fileNo
    Print #fileNo, "Diagnostic" & vbTab & "Severity" & vbTab & "Producer" & vbTab & "EventHistory" & vbTab & "Console" & vbTab & "ACK"
    CoreBridgeAnalyticsHarness_WriteMatrixRow fileNo, "DEADLINE_EXCEEDED", "WARNING", "CalcBridge_ComputeDeadlineAnalytics", CoreBridgeAnalyticsHarness_EventTypeExists("DEADLINE_EXCEEDED"), CoreBridgeAnalyticsHarness_CaptureHasText(capturePath, "Deadline exceeded"), True
    CoreBridgeAnalyticsHarness_WriteMatrixRow fileNo, "PARENT_DATES_IGNORED", "WARNING", "CalcBridge_ShowParentDateWarnings", CoreBridgeAnalyticsHarness_EventTypeExists("PARENT_DATES_IGNORED"), CoreBridgeAnalyticsHarness_CaptureHasText(capturePath, "Dates entered on summary task"), True
    CoreBridgeAnalyticsHarness_WriteMatrixRow fileNo, "TASK_TYPE_LOE_PROGRESS", "WARNING", "CalcBridge_AppendTaskTypeWarnings", CoreBridgeAnalyticsHarness_EventTypeExists("CONSOLE_WARNING"), CoreBridgeAnalyticsHarness_CaptureHasText(capturePath, "% Progress entered on LOE"), False
    CoreBridgeAnalyticsHarness_WriteMatrixRow fileNo, "ANALYTICS_TOPOLOGY_INCOMPLETE", "WARNING", "CalcBridge_AddAnalyticsTopologyWarning", CoreBridgeAnalyticsHarness_EventTypeExists("CONSOLE_WARNING"), CoreBridgeAnalyticsHarness_CaptureHasText(capturePath, "Analytics not calculated: incomplete topological order"), False
    Close #fileNo

End Sub

'------------------------------------------------------------------------------
' FR: Verifie le contrat Analytics Harness Write Matrix Row et signale toute divergence au harnais.
' EN: Verifies the Analytics Harness Write Matrix Row contract and reports any divergence to the harness.
'------------------------------------------------------------------------------

Private Sub CoreBridgeAnalyticsHarness_WriteMatrixRow( _
    ByVal fileNo As Integer, _
    ByVal diagnosticName As String, _
    ByVal severityName As String, _
    ByVal producerName As String, _
    ByVal eventHistoryPresent As Boolean, _
    ByVal consolePresent As Boolean, _
    ByVal ackPresent As Boolean)

    Print #fileNo, _
        CoreBridgeAnalyticsHarness_Tsv(diagnosticName) & vbTab & _
        CoreBridgeAnalyticsHarness_Tsv(severityName) & vbTab & _
        CoreBridgeAnalyticsHarness_Tsv(producerName) & vbTab & _
        CStr(eventHistoryPresent) & vbTab & _
        CStr(consolePresent) & vbTab & _
        CStr(ackPresent)

End Sub

'------------------------------------------------------------------------------
' FR: Verifie le contrat Analytics Harness Require Column et signale toute divergence au harnais.
' EN: Verifies the Analytics Harness Require Column contract and reports any divergence to the harness.
'------------------------------------------------------------------------------

Private Sub CoreBridgeAnalyticsHarness_RequireColumn(ByVal mapObj As Object, ByVal columnName As String)

    If mapObj Is Nothing Then Err.Raise vbObjectError + 9410, "CoreBridgeAnalyticsHarness", "Missing map"
    If Not mapObj.Exists(columnName) Then Err.Raise vbObjectError + 9411, "CoreBridgeAnalyticsHarness", "Missing column: " & columnName

End Sub

'------------------------------------------------------------------------------
' FR: Verifie le contrat Analytics Harness Event Type Exists et signale toute divergence au harnais.
' EN: Verifies the Analytics Harness Event Type Exists contract and reports any divergence to the harness.
'------------------------------------------------------------------------------

Private Function CoreBridgeAnalyticsHarness_EventTypeExists(ByVal eventType As String) As Boolean

    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim mapEvent As Object
    Dim arr As Variant
    Dim r As Long

    Set ws = ThisWorkbook.Worksheets("CALC_ALARM")
    Set tbl = ws.ListObjects("tbl_CALC_ALARM")
    If tbl.DataBodyRange Is Nothing Then Exit Function

    Set mapEvent = CanonicalIdentity_BuildColumnMap(tbl)
    If Not mapEvent.Exists("Event Type") Then Exit Function

    arr = tbl.DataBodyRange.Value
    For r = 1 To UBound(arr, 1)
        If UCase$(Trim$(CStr(arr(r, mapEvent("Event Type"))))) = UCase$(Trim$(eventType)) Then
            CoreBridgeAnalyticsHarness_EventTypeExists = True
            Exit Function
        End If
    Next r

End Function

'------------------------------------------------------------------------------
' FR: Verifie le contrat Analytics Harness Capture Has Text et signale toute divergence au harnais.
' EN: Verifies the Analytics Harness Capture Has Text contract and reports any divergence to the harness.
'------------------------------------------------------------------------------

Private Function CoreBridgeAnalyticsHarness_CaptureHasText( _
    ByVal capturePath As String, _
    ByVal expectedText As String) As Boolean

    Dim fileNo As Integer
    Dim lineText As String

    If Dir$(capturePath) = "" Then Exit Function

    fileNo = FreeFile
    Open capturePath For Input As #fileNo
    Do While Not EOF(fileNo)
        Line Input #fileNo, lineText
        If InStr(1, lineText, expectedText, vbTextCompare) > 0 Then
            CoreBridgeAnalyticsHarness_CaptureHasText = True
            Exit Do
        End If
    Loop
    Close #fileNo

End Function

'------------------------------------------------------------------------------
' FR: Verifie le contrat Analytics Harness Captured Message Count et signale toute divergence au harnais.
' EN: Verifies the Analytics Harness Captured Message Count contract and reports any divergence to the harness.
'------------------------------------------------------------------------------

Private Function CoreBridgeAnalyticsHarness_CapturedMessageCount(ByVal capturePath As String) As Long

    Dim fileNo As Integer
    Dim lineText As String
    Dim lineCount As Long

    If Dir$(capturePath) = "" Then Exit Function

    fileNo = FreeFile
    Open capturePath For Input As #fileNo
    Do While Not EOF(fileNo)
        Line Input #fileNo, lineText
        lineCount = lineCount + 1
    Loop
    Close #fileNo

    If lineCount > 1 Then CoreBridgeAnalyticsHarness_CapturedMessageCount = lineCount - 1

End Function

'------------------------------------------------------------------------------
' FR: Verifie le contrat Analytics Harness Assert et signale toute divergence au harnais.
' EN: Verifies the Analytics Harness Assert contract and reports any divergence to the harness.
'------------------------------------------------------------------------------

Private Sub CoreBridgeAnalyticsHarness_Assert(ByVal condition As Boolean, ByVal messageText As String)

    If Not condition Then
        CoreBridgeAnalyticsHarness_Trace "ASSERT FAIL " & messageText
        Err.Raise vbObjectError + 9412, "CoreBridgeAnalyticsHarness", messageText
    End If

End Sub

'------------------------------------------------------------------------------
' FR: Verifie le contrat Analytics Harness Trace et signale toute divergence au harnais.
' EN: Verifies the Analytics Harness Trace contract and reports any divergence to the harness.
'------------------------------------------------------------------------------

Private Sub CoreBridgeAnalyticsHarness_Trace(ByVal messageText As String)

    Dim fileNo As Integer

    On Error Resume Next
    If Trim$(gCoreBridgeAnalyticsTracePath) = "" Then Exit Sub
    fileNo = FreeFile
    Open gCoreBridgeAnalyticsTracePath For Append As #fileNo
    Print #fileNo, Format$(Now, "yyyy-mm-dd hh:nn:ss") & " | " & messageText
    Close #fileNo
    On Error GoTo 0

End Sub

'------------------------------------------------------------------------------
' FR: Verifie le contrat Analytics Harness Delete File et signale toute divergence au harnais.
' EN: Verifies the Analytics Harness Delete File contract and reports any divergence to the harness.
'------------------------------------------------------------------------------

Private Sub CoreBridgeAnalyticsHarness_DeleteFile(ByVal filePath As String)

    On Error Resume Next
    If Trim$(filePath) <> "" Then
        If Dir$(filePath) <> "" Then Kill filePath
    End If
    On Error GoTo 0

End Sub

'------------------------------------------------------------------------------
' FR: Verifie le contrat Analytics Harness Tsv et signale toute divergence au harnais.
' EN: Verifies the Analytics Harness Tsv contract and reports any divergence to the harness.
'------------------------------------------------------------------------------

Private Function CoreBridgeAnalyticsHarness_Tsv(ByVal value As String) As String

    CoreBridgeAnalyticsHarness_Tsv = Replace(Replace(Replace(value, vbTab, " "), vbCr, "\r"), vbLf, "\n")

End Function
