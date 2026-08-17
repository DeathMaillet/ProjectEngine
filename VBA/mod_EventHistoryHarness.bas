Attribute VB_Name = "mod_EventHistoryHarness"
Option Explicit

'===============================================================================
' MODULE : mod_EventHistoryHarness
' DOMAINE / DOMAIN : Validation Harnesses
'
' FR
' Harnais de preuve du contrat Event History sur des copies de test.
' N'appartient a aucun workflow produit et ne doit pas etre appele en usage normal.
'
' EN
' Proof harness for the Event History contract on test copies.
' Is not production workflow code and must not run during normal use.
'
' CONTRATS / CONTRACTS : EventHistoryHarness_Smoke
' CALLBACKS EXTERNES / EXTERNAL CALLBACKS : Aucun / None
'===============================================================================


Private gEventHistoryHarnessTracePath As String

'------------------------------------------------------------------------------
' FR: Execute un smoke deterministe du domaine EventHistory UI / ACK.
' EN: Runs a deterministic smoke for the EventHistory UI / ACK domain.
'------------------------------------------------------------------------------
Public Function EventHistoryHarness_Smoke(ByVal capturePath As String) As String

    Dim messages As Collection
    Dim displayMessages As Collection
    Dim warnItem As Object
    Dim eventHash As String
    Dim historyBefore As Long
    Dim historyAfter As Long
    Dim settingsInfoBefore As Boolean
    Dim formattedText As String

    On Error GoTo Fail

    gEventHistoryHarnessTracePath = capturePath & ".trace.txt"
    EventHistoryHarness_DeleteFile gEventHistoryHarnessTracePath
    EventHistoryHarness_Trace "01 enter"

    EnsurePlanningEventHistoryInfrastructure
    EventHistoryHarness_Trace "02 infrastructure ok"
    PlanningConsolePolicy_DisableNonInteractive
    EventHistory_SetLanguage "EN"
    EventHistory_SetShowInfo True
    ClearPlanningWarningAcknowledgements
    ClearPlanningEventHistory
    EventHistoryHarness_Trace "03 reset ok"

    eventHash = BuildPlanningEventHash( _
        "WARNING", "EH_HARNESS_WARNING", _
        "Avertissement EventHistory", "EventHistory warning", _
        "Detail FR", "Detail EN", _
        "EventHistoryHarness")

    LogPlanningEvent _
        "WARNING", "EH_HARNESS_WARNING", eventHash, _
        "Avertissement EventHistory", "EventHistory warning", _
        "Detail FR", "Detail EN", _
        "EventHistoryHarness", , , , "1.1", "Harness task", True
    EventHistoryHarness_Trace "04 log planning event ok"

    EventHistoryHarness_Assert EventHistoryHarness_TableRowCount("CALC_ALARM", "tbl_CALC_ALARM") > 0, "calc alarm event created"
    EventHistoryHarness_Assert EventHistoryHarness_TableRowCount("EVENT_HISTORY", "tbl_EVENT_HISTORY") > 0, "event history refreshed"
    EventHistoryHarness_Trace "05 event rows ok"

    Set messages = New Collection
    CalcBridge_AddConsoleMessage messages, "WARNING", _
        "FR:" & vbCrLf & "Avertissement EventHistory" & vbCrLf & vbCrLf & _
        "EN:" & vbCrLf & "EventHistory warning", _
        False, "EH_HARNESS_WARNING", eventHash
    Set warnItem = messages(1)
    EventHistoryHarness_Trace "06 warning message prepared"

    EventHistoryHarness_Assert PlanningMessage_CanAcknowledge(warnItem), "warning can acknowledge"
    EventHistoryHarness_Assert Not PlanningMessage_IsAcknowledged(warnItem), "warning starts unacknowledged"

    SetPlanningWarningAckState warnItem, True
    EventHistoryHarness_Trace "07 ack set returned"
    EventHistoryHarness_Assert PlanningMessage_IsAcknowledged(warnItem), "warning acknowledged"
    EventHistoryHarness_Assert EventHistoryHarness_TableRowCount("EVENT_ACK", "tbl_EVENT_ACK") > 0, "ack table populated"

    Set displayMessages = MessageEngine_PrepareDisplayMessages(messages)
    EventHistoryHarness_Assert displayMessages.Count = 0, "acknowledged warning hidden by MessageEngine"

    ClearPlanningWarningAcknowledgements
    EventHistoryHarness_Trace "08 ack clear returned"
    EventHistoryHarness_Assert EventHistoryHarness_TableRowCount("EVENT_ACK", "tbl_EVENT_ACK") = 0, "ack clear button contract clears table"
    EventHistoryHarness_Assert Not PlanningMessage_IsAcknowledged(warnItem), "warning unacknowledged after clear"

    EventHistory_SetShowInfo True
    Toggle_EventHistory_Info
    EventHistoryHarness_Trace "09 info toggle off returned"
    EventHistoryHarness_Assert Not EventHistory_CurrentShowInfo(), "show info toggle off"
    Toggle_EventHistory_Info
    EventHistoryHarness_Trace "10 info toggle on returned"
    EventHistoryHarness_Assert EventHistory_CurrentShowInfo(), "show info toggle on"

    EventHistory_SetLanguage "EN"
    formattedText = FormatPlanningConsoleMessageForCurrentLanguage( _
        "FR:" & vbCrLf & "Francais" & vbCrLf & vbCrLf & _
        "EN:" & vbCrLf & "English")
    EventHistoryHarness_Assert formattedText = "English", "language starts EN"
    Toggle_EventHistory_Language
    EventHistoryHarness_Trace "11 language toggle returned"
    formattedText = FormatPlanningConsoleMessageForCurrentLanguage( _
        "FR:" & vbCrLf & "Francais" & vbCrLf & vbCrLf & _
        "EN:" & vbCrLf & "English")
    EventHistoryHarness_Assert formattedText = "Francais", "language toggle switches to FR"
    EventHistory_SetLanguage "EN"

    Settings_Initialize
    EventHistoryHarness_Trace "12 settings initialize returned"
    settingsInfoBefore = EventHistory_CurrentShowInfo()
    Settings_ToggleInfoMessages
    EventHistoryHarness_Trace "13 settings info toggle returned"
    EventHistoryHarness_Assert EventHistory_CurrentShowInfo() <> settingsInfoBefore, "settings info toggle updates EventHistory"
    Settings_ToggleInfoMessages
    EventHistoryHarness_Trace "14 settings info restore returned"
    EventHistoryHarness_Assert EventHistory_CurrentShowInfo() = settingsInfoBefore, "settings info toggle restores EventHistory"

    EventHistory_SetLanguage "EN"
    ThisWorkbook.Worksheets("SETTINGS").Range("X8").value = "EN"
    Settings_ToggleEventHistoryLanguage
    EventHistoryHarness_Trace "15 settings event language toggle returned"
    formattedText = FormatPlanningConsoleMessageForCurrentLanguage( _
        "FR:" & vbCrLf & "Parametres FR" & vbCrLf & vbCrLf & _
        "EN:" & vbCrLf & "Settings EN")
    EventHistoryHarness_Assert formattedText = "Parametres FR", "settings language toggle updates EventHistory"
    EventHistory_SetLanguage "EN"

    EventHistoryHarness_Assert EventHistoryHarness_ShapeOnAction("EVENT_HISTORY", "btn_EventHistory_ClearHistory") = "ClearPlanningEventHistory", "clear history OnAction"
    EventHistoryHarness_Assert EventHistoryHarness_ShapeOnAction("EVENT_ACK", "btn_EventAck_ClearList") = "ClearPlanningWarningAcknowledgements", "clear ack OnAction"
    EventHistoryHarness_Trace "16 command OnAction ok"

    historyBefore = EventHistoryHarness_TableRowCount("EVENT_HISTORY", "tbl_EVENT_HISTORY")
    Set messages = New Collection
    CalcBridge_AddConsoleMessage messages, "INFO", _
        "FR:" & vbCrLf & "Info EventHistory" & vbCrLf & vbCrLf & _
        "EN:" & vbCrLf & "EventHistory info"
    PlanningConsolePolicy_EnableNonInteractive capturePath, "EventHistoryHarness"
    EventHistoryHarness_Trace "17 console policy enabled"
    CalcBridge_ShowPlanningConsole messages
    EventHistoryHarness_Trace "18 console show returned"
    EventHistoryHarness_Assert PlanningConsolePolicy_GetCapturedMessageCount() = 1, "planning console policy captures EventHistory info"
    PlanningConsolePolicy_DisableNonInteractive
    historyAfter = EventHistoryHarness_TableRowCount("EVENT_HISTORY", "tbl_EVENT_HISTORY")
    EventHistoryHarness_Assert historyAfter > historyBefore, "console policy interaction logs EventHistory"

    ClearPlanningEventHistory
    EventHistoryHarness_Trace "19 clear history returned"
    EventHistoryHarness_Assert EventHistoryHarness_TableRowCount("CALC_ALARM", "tbl_CALC_ALARM") = 0, "clear history clears calc alarm"
    EventHistoryHarness_Assert EventHistoryHarness_TableRowCount("EVENT_HISTORY", "tbl_EVENT_HISTORY") = 0, "clear history clears event history"

    EventHistory_SetLanguage "EN"
    EventHistory_SetShowInfo True
    PlanningConsolePolicy_DisableNonInteractive

    EventHistoryHarness_Trace "20 pass"
    EventHistoryHarness_Smoke = "PASS"
    Exit Function

Fail:
    On Error Resume Next
    EventHistoryHarness_Trace "FAIL " & Err.Number & " " & Err.Description
    PlanningConsolePolicy_DisableNonInteractive
    EventHistoryHarness_Smoke = "FAIL: " & Err.Description

End Function

'------------------------------------------------------------------------------
' FR: Verifie le contrat Event History Harness Assert et signale toute divergence au harnais.
' EN: Verifies the Event History Harness Assert contract and reports any divergence to the harness.
'------------------------------------------------------------------------------

Private Sub EventHistoryHarness_Assert(ByVal condition As Boolean, ByVal messageText As String)

    If Not condition Then
        EventHistoryHarness_Trace "ASSERT FAIL " & messageText
        Err.Raise vbObjectError + 9361, "EventHistoryHarness", messageText
    End If

End Sub

'------------------------------------------------------------------------------
' FR: Verifie le contrat Event History Harness Trace et signale toute divergence au harnais.
' EN: Verifies the Event History Harness Trace contract and reports any divergence to the harness.
'------------------------------------------------------------------------------

Private Sub EventHistoryHarness_Trace(ByVal messageText As String)

    Dim fileNo As Integer

    On Error Resume Next
    If Trim$(gEventHistoryHarnessTracePath) = "" Then Exit Sub
    fileNo = FreeFile
    Open gEventHistoryHarnessTracePath For Append As #fileNo
    Print #fileNo, Format$(Now, "yyyy-mm-dd hh:nn:ss") & " | " & messageText
    Close #fileNo

End Sub

'------------------------------------------------------------------------------
' FR: Verifie le contrat Event History Harness Delete File et signale toute divergence au harnais.
' EN: Verifies the Event History Harness Delete File contract and reports any divergence to the harness.
'------------------------------------------------------------------------------

Private Sub EventHistoryHarness_DeleteFile(ByVal filePath As String)

    On Error Resume Next
    If Len(Dir$(filePath)) > 0 Then Kill filePath

End Sub

'------------------------------------------------------------------------------
' FR: Verifie le contrat Event History Harness Table Row Count et signale toute divergence au harnais.
' EN: Verifies the Event History Harness Table Row Count contract and reports any divergence to the harness.
'------------------------------------------------------------------------------

Private Function EventHistoryHarness_TableRowCount( _
    ByVal sheetName As String, _
    ByVal tableName As String) As Long

    Dim tbl As ListObject

    On Error GoTo SafeExit
    Set tbl = ThisWorkbook.Worksheets(sheetName).ListObjects(tableName)
    If tbl.DataBodyRange Is Nothing Then Exit Function
    EventHistoryHarness_TableRowCount = tbl.ListRows.Count

SafeExit:
End Function

'------------------------------------------------------------------------------
' FR: Verifie le contrat Event History Harness Shape On Action et signale toute divergence au harnais.
' EN: Verifies the Event History Harness Shape On Action contract and reports any divergence to the harness.
' FR - Effet de bord : cree ou met a jour des shapes Excel.
' EN - Side effect: creates or updates Excel shapes.
'------------------------------------------------------------------------------

Private Function EventHistoryHarness_ShapeOnAction( _
    ByVal sheetName As String, _
    ByVal shapeName As String) As String

    On Error GoTo SafeExit
    EventHistoryHarness_ShapeOnAction = CStr(ThisWorkbook.Worksheets(sheetName).Shapes(shapeName).OnAction)

SafeExit:
End Function
