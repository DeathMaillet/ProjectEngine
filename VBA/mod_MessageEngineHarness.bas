Attribute VB_Name = "mod_MessageEngineHarness"
Option Explicit

'===============================================================================
' MODULE : mod_MessageEngineHarness
' DOMAINE / DOMAIN : Validation Harnesses
'
' FR
' Harnais de preuve du contrat Message Engine sur des copies de test.
' N'appartient a aucun workflow produit et ne doit pas etre appele en usage normal.
'
' EN
' Proof harness for the Message Engine contract on test copies.
' Is not production workflow code and must not run during normal use.
'
' CONTRATS / CONTRACTS : MessageEngineHarness_Smoke
' CALLBACKS EXTERNES / EXTERNAL CALLBACKS : Aucun / None
'===============================================================================


'------------------------------------------------------------------------------
' FR: Execute un smoke deterministe du contrat MessageEngine / EventHistory.
' EN: Runs a deterministic smoke for the MessageEngine / EventHistory contract.
'------------------------------------------------------------------------------
Public Function MessageEngineHarness_Smoke(ByVal capturePath As String) As String

    Dim messages As Collection
    Dim consoleMessages As Collection
    Dim displayMessages As Collection
    Dim warnItem As Object
    Dim stopGrouped As Collection
    Dim historyBefore As Long
    Dim historyAfter As Long

    On Error GoTo Fail

    EnsurePlanningEventHistoryInfrastructure
    ClearPlanningWarningAcknowledgements
    ClearPlanningEventHistory
    EventHistory_SetShowInfo True
    PlanningConsolePolicy_DisableNonInteractive

    Set messages = New Collection
    CalcBridge_AddConsoleMessage messages, "INFO", "FR:" & vbCrLf & "Info FR" & vbCrLf & vbCrLf & "EN:" & vbCrLf & "Info EN"
    CalcBridge_AddConsoleMessage messages, "WARNING", "FR:" & vbCrLf & "Warning FR" & vbCrLf & vbCrLf & "EN:" & vbCrLf & "Warning EN", False, "HARNESS_WARNING", "HARNESS-HASH-001"
    CalcBridge_AddConsoleMessage messages, "STOP", "FR:" & vbCrLf & "Stop FR" & vbCrLf & vbCrLf & "EN:" & vbCrLf & "Stop EN"

    Set consoleMessages = MessageEngine_PrepareConsoleMessages(messages)
    MessageEngineHarness_Assert consoleMessages.Count = 3, "console message count"
    MessageEngineHarness_Assert CStr(consoleMessages(1)("Type")) = "STOP", "console order stop first"
    MessageEngineHarness_Assert CStr(consoleMessages(2)("Type")) = "WARNING", "console order warning second"
    MessageEngineHarness_Assert CStr(consoleMessages(3)("Type")) = "INFO", "console order info third"

    Set displayMessages = MessageEngine_PrepareDisplayMessages(consoleMessages)
    MessageEngineHarness_Assert MessageEngine_ShouldShowConsole(displayMessages), "display should show"
    MessageEngineHarness_Assert displayMessages.Count = 3, "display message count before ack"

    Set warnItem = consoleMessages(2)
    MessageEngineHarness_Assert PlanningMessage_CanAcknowledge(warnItem), "warning can acknowledge"
    MessageEngineHarness_Assert Not PlanningMessage_IsAcknowledged(warnItem), "warning starts unacknowledged"
    SetPlanningWarningAckState warnItem, True
    MessageEngineHarness_Assert PlanningMessage_IsAcknowledged(warnItem), "warning acknowledged"

    Set displayMessages = MessageEngine_PrepareDisplayMessages(consoleMessages)
    MessageEngineHarness_Assert displayMessages.Count = 2, "acknowledged warning hidden from display"
    SetPlanningWarningAckState warnItem, False
    MessageEngineHarness_Assert Not PlanningMessage_IsAcknowledged(warnItem), "warning unacknowledged again"

    Set stopGrouped = New Collection
    CalcBridge_AddConsoleMessage stopGrouped, "STOP", "FR:" & vbCrLf & "Stop A FR" & vbCrLf & vbCrLf & "EN:" & vbCrLf & "Stop A EN"
    CalcBridge_AddConsoleMessage stopGrouped, "STOP", "FR:" & vbCrLf & "Stop B FR" & vbCrLf & vbCrLf & "EN:" & vbCrLf & "Stop B EN"
    Set displayMessages = MessageEngine_PrepareDisplayMessages(stopGrouped)
    MessageEngineHarness_Assert displayMessages.Count = 1, "multiple stops grouped"
    MessageEngineHarness_Assert InStr(1, CStr(displayMessages(1)("Message")), String$(36, "-"), vbBinaryCompare) > 0, "grouped stop separator"

    EventHistory_SetShowInfo False
    Set displayMessages = MessageEngine_PrepareDisplayMessages(consoleMessages)
    MessageEngineHarness_Assert displayMessages.Count = 2, "info hidden when info toggle off"
    EventHistory_SetShowInfo True

    historyBefore = MessageEngineHarness_TableRowCount("EVENT_HISTORY", "tbl_EVENT_HISTORY")
    PlanningConsolePolicy_EnableNonInteractive capturePath, "MessageEngineHarness"
    CalcBridge_ShowPlanningConsole messages
    MessageEngineHarness_Assert PlanningConsolePolicy_GetCapturedMessageCount() = 3, "noninteractive capture count"
    PlanningConsolePolicy_DisableNonInteractive
    historyAfter = MessageEngineHarness_TableRowCount("EVENT_HISTORY", "tbl_EVENT_HISTORY")
    MessageEngineHarness_Assert historyAfter > historyBefore, "event history logged console messages"

    MessageEngineHarness_Smoke = "PASS"
    Exit Function

Fail:
    On Error Resume Next
    PlanningConsolePolicy_DisableNonInteractive
    MessageEngineHarness_Smoke = "FAIL: " & Err.Description
    Err.Raise vbObjectError + 9315, "MessageEngineHarness_Smoke", MessageEngineHarness_Smoke

End Function

'------------------------------------------------------------------------------
' FR: Verifie le contrat Message Engine Harness Assert et signale toute divergence au harnais.
' EN: Verifies the Message Engine Harness Assert contract and reports any divergence to the harness.
'------------------------------------------------------------------------------

Private Sub MessageEngineHarness_Assert(ByVal condition As Boolean, ByVal messageText As String)

    If Not condition Then Err.Raise vbObjectError + 9316, "MessageEngineHarness", messageText

End Sub

'------------------------------------------------------------------------------
' FR: Verifie le contrat Message Engine Harness Table Row Count et signale toute divergence au harnais.
' EN: Verifies the Message Engine Harness Table Row Count contract and reports any divergence to the harness.
'------------------------------------------------------------------------------

Private Function MessageEngineHarness_TableRowCount( _
    ByVal sheetName As String, _
    ByVal tableName As String) As Long

    Dim tbl As ListObject

    On Error GoTo SafeExit
    Set tbl = ThisWorkbook.Worksheets(sheetName).ListObjects(tableName)
    If tbl.DataBodyRange Is Nothing Then Exit Function
    MessageEngineHarness_TableRowCount = tbl.ListRows.Count

SafeExit:
End Function
