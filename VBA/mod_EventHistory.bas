Attribute VB_Name = "mod_EventHistory"
Option Explicit

'===============================================================================
' MODULE : mod_EventHistory
' DOMAINE / DOMAIN : Event History / ACK
'
' FR
' Possede le journal des evenements planning, les ACK et leur projection UI.
' Ne produit pas les diagnostics metier qu'il journalise.
'
' EN
' Owns planning event history, ACK state and their UI projection.
' Does not produce the business diagnostics it records.
'
' CONTRATS / CONTRACTS : BeginPlanningEventRun, EndPlanningEventRun, EventHistory_SetLanguage, EventHistory_ApplyLanguage, ShouldShowInfoOnlyPlanningConsole, EventHistory_SetShowInfo, EventHistory_CurrentShowInfo, Toggle_EventHistory_Info
' CALLBACKS EXTERNES / EXTERNAL CALLBACKS : Aucun / None
'===============================================================================


Private Const CALC_ALARM_SHEET As String = "CALC_ALARM"
Private Const CALC_ALARM_TABLE As String = "tbl_CALC_ALARM"
Private Const EVENT_HISTORY_SHEET As String = "EVENT_HISTORY"
Private Const EVENT_HISTORY_TABLE As String = "tbl_EVENT_HISTORY"
Private Const EVENT_ACK_SHEET As String = "EVENT_ACK"
Private Const EVENT_ACK_TABLE As String = "tbl_EVENT_ACK"
Private Const EVENT_HISTORY_HEADER_ROW As Long = 3
Private Const EVENT_ACK_HEADER_ROW As Long = 2

Private Const EVENT_HISTORY_INFO_LABEL As String = "btn_EventHistory_Info_Label"
Private Const EVENT_HISTORY_INFO_BG As String = "btn_EventHistory_Info_BG"
Private Const EVENT_HISTORY_INFO_KNOB As String = "btn_EventHistory_Info_Knob"
Private Const EVENT_HISTORY_LANG_LABEL As String = "btn_EventHistory_Language_Label"
Private Const EVENT_HISTORY_LANG_BG As String = "btn_EventHistory_Language_BG"
Private Const EVENT_HISTORY_LANG_KNOB As String = "btn_EventHistory_Language_Knob"
Private Const EVENT_ACK_CLEAR_BUTTON As String = "btn_EventAck_ClearList"
Private Const EVENT_HISTORY_CLEAR_BUTTON As String = "btn_EventHistory_ClearHistory"

Private gPlanningEventRunId As String
Private gPlanningEventRunActive As Boolean
Private gEventHistoryShowInfo As Boolean
Private gEventHistoryShowInfoInitialized As Boolean
Private gPlanningLanguage As String
Private gEventHistoryInternalWriteDepth As Long

'------------------------------------------------------------------------------
' FR: Ouvre le cycle de traitement Planning Event Run.
' EN: Begins the Planning Event Run processing cycle.
'------------------------------------------------------------------------------
Public Sub BeginPlanningEventRun(Optional ByVal sourceProcedure As String = "")

    gPlanningEventRunId = BuildPlanningEventRunId(sourceProcedure)
    gPlanningEventRunActive = True

End Sub

'------------------------------------------------------------------------------
' FR: Ferme le cycle de traitement Planning Event Run.
' EN: Ends the Planning Event Run processing cycle.
'------------------------------------------------------------------------------
Public Sub EndPlanningEventRun()

    gPlanningEventRunActive = False
    gPlanningEventRunId = vbNullString

End Sub

'------------------------------------------------------------------------------
' FR: Active ou initialise Set Language dans l'etat runtime du composant.
' EN: Activates or initializes Set Language in the component runtime state.
'------------------------------------------------------------------------------

Public Sub EventHistory_SetLanguage(ByVal languageCode As String)

    EnsureEventHistoryState

    Select Case UCase$(Trim$(languageCode))
        Case "EN"
            gPlanningLanguage = "EN"
        Case "FR"
            gPlanningLanguage = "FR"
        Case Else
            gPlanningLanguage = "FR"
    End Select

End Sub

'------------------------------------------------------------------------------
' FR: Actualise Apply Language sans modifier les regles metier qui produisent les donnees.
' EN: Refreshes Apply Language without changing the business rules that produce the data.
'------------------------------------------------------------------------------

Public Sub EventHistory_ApplyLanguage(Optional ByVal languageCode As String = "")

    If Trim$(languageCode) <> "" Then EventHistory_SetLanguage languageCode
    Refresh_EventHistory_View False, True

End Sub

'------------------------------------------------------------------------------
' FR: Retourne la valeur Language sans exposer de mutateur sur l'etat source.
' EN: Returns the Language value without exposing a mutator for source state.
'------------------------------------------------------------------------------

Public Function EventHistory_CurrentLanguage() As String

    EnsureEventHistoryState
    EventHistory_CurrentLanguage = gPlanningLanguage

End Function

'------------------------------------------------------------------------------
' FR: Retourne Planning Console Language depuis le contexte event history and acknowledgements.
' EN: Returns Planning Console Language from the event history and acknowledgements context.
'------------------------------------------------------------------------------
Private Function GetPlanningConsoleLanguage() As String

    GetPlanningConsoleLanguage = EventHistory_CurrentLanguage()

End Function

'------------------------------------------------------------------------------
' FR: Indique si Should Show Info Only Planning Console est vrai pour le contexte courant.
' EN: Returns whether Should Show Info Only Planning Console is true for the current context.
'------------------------------------------------------------------------------
Public Function ShouldShowInfoOnlyPlanningConsole() As Boolean

    EnsureEventHistoryState
    ShouldShowInfoOnlyPlanningConsole = gEventHistoryShowInfo

End Function

'------------------------------------------------------------------------------
' FR: Active ou initialise Set Show Info dans l'etat runtime du composant.
' EN: Activates or initializes Set Show Info in the component runtime state.
'------------------------------------------------------------------------------

Public Sub EventHistory_SetShowInfo(ByVal showInfo As Boolean)

    EnsureEventHistoryState
    gEventHistoryShowInfo = showInfo

End Sub

'------------------------------------------------------------------------------
' FR: Retourne la valeur Show Info sans exposer de mutateur sur l'etat source.
' EN: Returns the Show Info value without exposing a mutator for source state.
'------------------------------------------------------------------------------

Public Function EventHistory_CurrentShowInfo() As Boolean

    EventHistory_CurrentShowInfo = ShouldShowInfoOnlyPlanningConsole()

End Function

'------------------------------------------------------------------------------
' FR: Bascule l'etat Event History Info et met a jour les sorties associees.
' EN: Toggles Event History Info state and updates related outputs.
'------------------------------------------------------------------------------
Public Sub Toggle_EventHistory_Info()
    Settings_ToggleInfoMessages

End Sub

'------------------------------------------------------------------------------
' FR: Bascule l'etat Event History Language et met a jour les sorties associees.
' EN: Toggles Event History Language state and updates related outputs.
'------------------------------------------------------------------------------
Public Sub Toggle_EventHistory_Language()
    Settings_ToggleEventHistoryLanguage

End Sub

'------------------------------------------------------------------------------
' FR: Traite un changement ou evenement pour Event History Change.
' EN: Handles a change or event for Event History Change.
'------------------------------------------------------------------------------
Public Sub Handle_EventHistory_Change(ByVal ws As Worksheet, ByVal Target As Range)

    On Error GoTo SafeExit

    If IsPlanningEventInternalWriteActive() Then Exit Sub

    If Not ws Is Nothing Then
        If CStr(ws.Name) = EVENT_ACK_SHEET Then
            If IsAllowedManualEventAckChange(ws, Target) Then
                Application.EnableEvents = False
                NormalizeManualEventAckRows ws, Target
                Refresh_EventHistory_View
                GoTo SafeExit
            End If
        End If
    End If

    Application.EnableEvents = False
    Application.Undo

SafeExit:
    If Not IsMacroAbortRequested() Then
        Application.EnableEvents = True
    End If

End Sub

'------------------------------------------------------------------------------
' FR: Indique si Allowed Manual Event Ack Change est vrai pour le contexte courant.
' EN: Returns whether Allowed Manual Event Ack Change is true for the current context.
'------------------------------------------------------------------------------
Private Function IsAllowedManualEventAckChange(ByVal ws As Worksheet, ByVal Target As Range) As Boolean

    Dim tblAck As ListObject
    Dim editableRange As Range

    If ws Is Nothing Then Exit Function
    If Target Is Nothing Then Exit Function

    On Error Resume Next
    Set tblAck = ws.ListObjects(EVENT_ACK_TABLE)
    On Error GoTo 0

    If tblAck Is Nothing Then Exit Function
    If Intersect(Target, tblAck.HeaderRowRange) Is Nothing Then
        If Not tblAck.DataBodyRange Is Nothing Then
            Set editableRange = tblAck.DataBodyRange
        Else
            Set editableRange = tblAck.Range.Offset(1, 0).Resize(1, tblAck.Range.Columns.Count)
        End If

        IsAllowedManualEventAckChange = Not Intersect(Target, editableRange) Is Nothing
    End If

End Function

'------------------------------------------------------------------------------
' FR: Normalise Manual Event Ack Rows dans un format exploitable.
' EN: Normalizes Manual Event Ack Rows into a usable format.
'------------------------------------------------------------------------------
Private Sub NormalizeManualEventAckRows(ByVal ws As Worksheet, ByVal Target As Range)

    Dim tblAck As ListObject
    Dim touched As Range
    Dim rowRange As Range
    Dim acknowledgedCell As Range
    Dim severityCell As Range

    If ws Is Nothing Then Exit Sub
    If Target Is Nothing Then Exit Sub

    On Error Resume Next
    Set tblAck = ws.ListObjects(EVENT_ACK_TABLE)
    On Error GoTo 0

    If tblAck Is Nothing Then Exit Sub
    If tblAck.DataBodyRange Is Nothing Then Exit Sub

    Set touched = Intersect(Target, tblAck.DataBodyRange)
    If touched Is Nothing Then Exit Sub

    For Each rowRange In touched.rows
        Set acknowledgedCell = Intersect(rowRange.EntireRow, SchemaListColumn(tblAck, VTS_TABLE_EVENT_ACK, VTS_COL_ACKNOWLEDGED).DataBodyRange)
        If Not acknowledgedCell Is Nothing Then
            If Trim$(CStr(acknowledgedCell.Cells(1, 1).value)) = "" Then acknowledgedCell.Cells(1, 1).value = True
        End If

        Set severityCell = Intersect(rowRange.EntireRow, SchemaListColumn(tblAck, VTS_TABLE_EVENT_ACK, VTS_COL_SEVERITY).DataBodyRange)
        If Not severityCell Is Nothing Then
            If Trim$(CStr(severityCell.Cells(1, 1).value)) <> "" Then severityCell.Cells(1, 1).value = UCase$(Trim$(CStr(severityCell.Cells(1, 1).value)))
        End If
    Next rowRange

End Sub


'------------------------------------------------------------------------------
' FR: Rafraichit Event History View a partir de l'etat courant.
' EN: Refreshes Event History View from the current state.
'------------------------------------------------------------------------------
Public Sub Refresh_EventHistory_View( _
    Optional ByVal hydratePersistedLanguage As Boolean = True, _
    Optional ByVal propagateErrors As Boolean = False)

    Dim perfScope As clsPerfScope

    Dim wsAlarm As Worksheet
    Dim wsHistory As Worksheet
    Dim wsAck As Worksheet
    Dim tblAlarm As ListObject
    Dim tblHistory As ListObject
    Dim tblAck As ListObject
    Dim ackLookup As Object

    Dim r As Long
    Dim newRow As ListRow
    Dim msgText As String
    Dim severity As String
    Dim eventType As String
    Dim eventHash As String
    Dim internalWriteStarted As Boolean
    Dim oldScreenUpdating As Boolean
    Dim errorNumber As Long
    Dim errorSource As String
    Dim errorDescription As String

    Set perfScope = Profiler_BeginScope("Refresh_EventHistory_View", "Event History")
    If hydratePersistedLanguage Then _
        EventHistory_SetLanguage Settings_GetOwnerLanguage("EVENT")

    oldScreenUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False
    On Error GoTo CleanFail
    BeginPlanningEventInternalWrite wsAlarm, wsHistory, wsAck
    internalWriteStarted = True
    Set tblAlarm = wsAlarm.ListObjects(CALC_ALARM_TABLE)
    Set tblHistory = wsHistory.ListObjects(EVENT_HISTORY_TABLE)
    Set tblAck = wsAck.ListObjects(EVENT_ACK_TABLE)
    Set ackLookup = BuildEventAckLookup(tblAck)

    ClearPlanningEventTableRows tblHistory

    If Not tblAlarm.DataBodyRange Is Nothing Then
        For r = tblAlarm.ListRows.Count To 1 Step -1
            severity = UCase$(Trim$(CStr(tblAlarm.DataBodyRange.Cells(r, tblAlarm.ListColumns("Severity").Index).value)))
            eventType = Trim$(CStr(tblAlarm.DataBodyRange.Cells(r, tblAlarm.ListColumns("Event Type").Index).value))
            eventHash = Trim$(CStr(tblAlarm.DataBodyRange.Cells(r, tblAlarm.ListColumns("Hash").Index).value))
            msgText = BuildEventHistoryDisplayMessage( _
                CStr(tblAlarm.DataBodyRange.Cells(r, tblAlarm.ListColumns("FR Message").Index).value), _
                CStr(tblAlarm.DataBodyRange.Cells(r, tblAlarm.ListColumns("EN Message").Index).value), _
                CStr(tblAlarm.DataBodyRange.Cells(r, tblAlarm.ListColumns("FR Details").Index).value), _
                CStr(tblAlarm.DataBodyRange.Cells(r, tblAlarm.ListColumns("EN Details").Index).value))

            Set newRow = tblHistory.ListRows.Add
            With newRow.Range
                .Cells(1, SchemaListColumn(tblHistory, VTS_TABLE_EVENT_HISTORY, VTS_COL_DATE).Index).value = tblAlarm.DataBodyRange.Cells(r, tblAlarm.ListColumns("Date").Index).value
                .Cells(1, SchemaListColumn(tblHistory, VTS_TABLE_EVENT_HISTORY, VTS_COL_HOUR).Index).value = tblAlarm.DataBodyRange.Cells(r, tblAlarm.ListColumns("Time").Index).value
                .Cells(1, SchemaListColumn(tblHistory, VTS_TABLE_EVENT_HISTORY, VTS_COL_SEVERITY).Index).value = severity
                .Cells(1, SchemaListColumn(tblHistory, VTS_TABLE_EVENT_HISTORY, VTS_COL_MESSAGE).Index).value = msgText
                .Cells(1, SchemaListColumn(tblHistory, VTS_TABLE_EVENT_HISTORY, VTS_COL_ACKNOWLEDGED).Index).value = IsEventAcknowledged(ackLookup, severity, eventType, eventHash)
            End With
        Next r
    End If

    ApplyPlanningEventFormats tblAlarm, tblHistory, tblAck
    EnsureEventHistoryToggleShapes wsHistory
    RefreshEventHistoryToggleVisuals wsHistory

CleanExit:
    If internalWriteStarted Then
        EndPlanningEventInternalWrite wsAlarm, wsHistory, wsAck
    End If
    Application.ScreenUpdating = oldScreenUpdating
    If errorNumber <> 0 And propagateErrors Then _
        Err.Raise errorNumber, errorSource, errorDescription
    Exit Sub

CleanFail:
    errorNumber = Err.Number
    errorSource = Err.Source
    errorDescription = Err.Description
    Resume CleanExit

End Sub

'------------------------------------------------------------------------------
' FR: Formate Planning Console Message For Current Language pour l'affichage ou l'ecriture.
' EN: Formats Planning Console Message For Current Language for display or writing.
'------------------------------------------------------------------------------
Public Function FormatPlanningConsoleMessageForCurrentLanguage(ByVal rawMessage As String) As String

    Dim txt As String
    Dim upperTxt As String
    Dim enPos As Long
    Dim enMarkerLen As Long
    Dim frPart As String
    Dim enPart As String

    EnsureEventHistoryState

    txt = TrimPlanningMessageLineEdges(CStr(rawMessage))
    upperTxt = UCase$(txt)

    If Left$(upperTxt, 3) <> "FR:" Then
        FormatPlanningConsoleMessageForCurrentLanguage = TrimPlanningMessageLineEdges(txt)
        Exit Function
    End If

    enPos = InStr(1, upperTxt, vbCrLf & "EN:", vbTextCompare)
    enMarkerLen = Len(vbCrLf & "EN:")
    If enPos = 0 Then
        enPos = InStr(1, upperTxt, vbLf & "EN:", vbTextCompare)
        enMarkerLen = Len(vbLf & "EN:")
    End If
    If enPos = 0 Then
        FormatPlanningConsoleMessageForCurrentLanguage = TrimPlanningMessageLineEdges(txt)
        Exit Function
    End If

    frPart = Mid$(txt, 4, enPos - 4)
    enPart = Mid$(txt, enPos + enMarkerLen)

    If UCase$(Trim$(gPlanningLanguage)) = "EN" Then
        FormatPlanningConsoleMessageForCurrentLanguage = TrimPlanningMessageLineEdges(enPart)
    Else
        FormatPlanningConsoleMessageForCurrentLanguage = TrimPlanningMessageLineEdges(frPart)
    End If

End Function

'------------------------------------------------------------------------------
' FR: Construit la valeur Planning Warning Ack Token a partir des donnees fournies par l'appelant.
' EN: Builds the Planning Warning Ack Token value from data supplied by the caller.
'------------------------------------------------------------------------------

Public Function BuildPlanningWarningAckToken( _
    ByVal eventType As String, _
    ByVal eventHash As String) As String

    If Trim$(eventType) = "" Then Exit Function
    If Trim$(eventHash) = "" Then Exit Function

    BuildPlanningWarningAckToken = UCase$(Trim$(eventType)) & "|" & Trim$(eventHash)

End Function

'------------------------------------------------------------------------------
' FR: Retourne la map Planning Message Is Acknowledged sans modifier les donnees d'entree.
' EN: Returns the Planning Message Is Acknowledged map without mutating input data.
'------------------------------------------------------------------------------

Public Function PlanningMessage_IsAcknowledged(ByVal item As Variant) As Boolean

    Dim msgType As String
    Dim ackTokens As String
    Dim tblAck As ListObject
    Dim ackLookup As Object

    On Error GoTo SafeExit

    msgType = NormalizeConsoleEventSeverity(CStr(item("Type")))
    If msgType <> "WARNING" Then Exit Function

    On Error Resume Next
    Set tblAck = ThisWorkbook.Worksheets(EVENT_ACK_SHEET).ListObjects(EVENT_ACK_TABLE)
    On Error GoTo SafeExit
    If tblAck Is Nothing Then
        EnsurePlanningEventHistoryInfrastructure
        Set tblAck = ThisWorkbook.Worksheets(EVENT_ACK_SHEET).ListObjects(EVENT_ACK_TABLE)
    End If

    Set ackLookup = BuildEventAckLookup(tblAck)

    ackTokens = PlanningMessage_BuildAckTokens(item)
    If Trim$(ackTokens) <> "" Then
        PlanningMessage_IsAcknowledged = ArePlanningWarningAckTokensAcknowledged(ackLookup, ackTokens)
        Exit Function
    End If

SafeExit:
End Function

'------------------------------------------------------------------------------
' FR: Retourne la valeur Planning Message Can Acknowledge sans modifier les donnees d'entree.
' EN: Returns the Planning Message Can Acknowledge value without mutating input data.
'------------------------------------------------------------------------------

Public Function PlanningMessage_CanAcknowledge(ByVal item As Variant) As Boolean

    On Error GoTo SafeExit

    If NormalizeConsoleEventSeverity(CStr(item("Type"))) <> "WARNING" Then Exit Function
    PlanningMessage_CanAcknowledge = (Trim$(PlanningMessage_BuildAckTokens(item)) <> "")

SafeExit:
End Function

'------------------------------------------------------------------------------
' FR: Met a jour Planning Warning Ack State dans le contexte event history and acknowledgements.
' EN: Updates Planning Warning Ack State in the event history and acknowledgements context.
'------------------------------------------------------------------------------
Public Sub SetPlanningWarningAckState( _
    ByVal item As Variant, _
    ByVal acknowledged As Boolean)

    Dim ackTokens As String
    Dim tokens() As String
    Dim oneToken As Variant
    Dim wsAlarm As Worksheet
    Dim wsHistory As Worksheet
    Dim wsAck As Worksheet
    Dim internalWriteStarted As Boolean

    On Error GoTo SafeExit

    If Not PlanningMessage_CanAcknowledge(item) Then Exit Sub

    BeginPlanningEventInternalWrite wsAlarm, wsHistory, wsAck
    internalWriteStarted = True

    ackTokens = PlanningMessage_BuildAckTokens(item)
    tokens = Split(ackTokens, ";")

    For Each oneToken In tokens
        If Trim$(CStr(oneToken)) <> "" Then
            If acknowledged Then
                UpsertPlanningWarningAckToken CStr(oneToken), item
            Else
                RemovePlanningWarningAckToken CStr(oneToken)
            End If
        End If
    Next oneToken

    item("Acknowledged") = PlanningMessage_IsAcknowledged(item)
    Refresh_EventHistory_View

SafeExit:
    If internalWriteStarted Then
        EndPlanningEventInternalWrite wsAlarm, wsHistory, wsAck
    End If
End Sub

'------------------------------------------------------------------------------
' FR: Vide ou reinitialise Planning Warning Acknowledgements.
' EN: Clears or resets Planning Warning Acknowledgements.
'------------------------------------------------------------------------------
Public Sub ClearPlanningWarningAcknowledgements()

    Dim wsAlarm As Worksheet
    Dim wsHistory As Worksheet
    Dim wsAck As Worksheet
    Dim tblAck As ListObject
    Dim internalWriteStarted As Boolean

    On Error GoTo CleanFail
    BeginPlanningEventInternalWrite wsAlarm, wsHistory, wsAck
    internalWriteStarted = True
    Set tblAck = wsAck.ListObjects(EVENT_ACK_TABLE)

    ClearPlanningEventTableRows tblAck

    Refresh_EventHistory_View

CleanExit:
    If internalWriteStarted Then
        EndPlanningEventInternalWrite wsAlarm, wsHistory, wsAck
    End If
    Exit Sub

CleanFail:
    Resume CleanExit

End Sub

'------------------------------------------------------------------------------
' FR: Vide ou reinitialise Planning Event History.
' EN: Clears or resets Planning Event History.
'------------------------------------------------------------------------------
Public Sub ClearPlanningEventHistory()

    Dim wsAlarm As Worksheet
    Dim wsHistory As Worksheet
    Dim wsAck As Worksheet
    Dim tblAlarm As ListObject
    Dim tblHistory As ListObject
    Dim internalWriteStarted As Boolean

    On Error GoTo CleanFail
    BeginPlanningEventInternalWrite wsAlarm, wsHistory, wsAck
    internalWriteStarted = True
    Set tblAlarm = wsAlarm.ListObjects(CALC_ALARM_TABLE)
    Set tblHistory = wsHistory.ListObjects(EVENT_HISTORY_TABLE)

    ClearPlanningEventTableRows tblAlarm
    ClearPlanningEventTableRows tblHistory

    Refresh_EventHistory_View

CleanExit:
    If internalWriteStarted Then
        EndPlanningEventInternalWrite wsAlarm, wsHistory, wsAck
    End If
    Exit Sub

CleanFail:
    Resume CleanExit

End Sub

'------------------------------------------------------------------------------
' FR: Construit la valeur Planning Message Build Ack Tokens a partir des donnees fournies par l'appelant.
' EN: Builds the Planning Message Build Ack Tokens value from data supplied by the caller.
'------------------------------------------------------------------------------

Private Function PlanningMessage_BuildAckTokens(ByVal item As Variant) As String

    Dim rawMessage As String
    Dim eventType As String
    Dim eventHash As String
    Dim frMessage As String
    Dim enMessage As String
    Dim frDetails As String
    Dim enDetails As String

    On Error GoTo SafeExit

    If NormalizeConsoleEventSeverity(CStr(item("Type"))) <> "WARNING" Then Exit Function

    PlanningMessage_BuildAckTokens = PlanningMessage_GetOptionalText(item, "AckTokens")
    If Trim$(PlanningMessage_BuildAckTokens) <> "" Then Exit Function

    eventType = PlanningMessage_GetOptionalText(item, "EventType")
    eventHash = PlanningMessage_GetOptionalText(item, "Hash")

    If eventType = "" Or eventHash = "" Then
        rawMessage = CStr(item("Message"))
        SplitConsoleMessageForHistory rawMessage, frMessage, enMessage, frDetails, enDetails
        eventType = "CONSOLE_WARNING"
        eventHash = BuildPlanningEventHash( _
            "WARNING", eventType, frMessage, enMessage, frDetails, enDetails, _
            "MessageEngine")
    End If

    PlanningMessage_BuildAckTokens = BuildPlanningWarningAckToken(eventType, eventHash)

SafeExit:
End Function

'------------------------------------------------------------------------------
' FR: Verifie ou cree Planning Event History Infrastructure si necessaire.
' EN: Ensures or creates Planning Event History Infrastructure when needed.
'------------------------------------------------------------------------------
Public Sub EnsurePlanningEventHistoryInfrastructure()

    Dim wsAlarm As Worksheet
    Dim wsHistory As Worksheet
    Dim wsAck As Worksheet
    Dim tblAlarm As ListObject
    Dim tblHistory As ListObject
    Dim tblAck As ListObject
    Dim internalWriteStarted As Boolean

    On Error GoTo CleanFail
    gEventHistoryInternalWriteDepth = gEventHistoryInternalWriteDepth + 1
    internalWriteStarted = True

    Set wsAlarm = EnsurePlanningEventSheet(CALC_ALARM_SHEET)
    Set wsHistory = EnsurePlanningEventSheet(EVENT_HISTORY_SHEET)
    Set wsAck = EnsurePlanningEventSheet(EVENT_ACK_SHEET)

    EnsurePlanningEventTable wsAlarm, CALC_ALARM_TABLE, CalcAlarmHeaders()
    EnsurePlanningEventTable wsHistory, EVENT_HISTORY_TABLE, EventHistoryHeaders(), EVENT_HISTORY_HEADER_ROW
    EnsurePlanningEventTable wsAck, EVENT_ACK_TABLE, EventAckHeaders(), EVENT_ACK_HEADER_ROW

    Set tblAlarm = wsAlarm.ListObjects(CALC_ALARM_TABLE)
    Set tblHistory = wsHistory.ListObjects(EVENT_HISTORY_TABLE)
    Set tblAck = wsAck.ListObjects(EVENT_ACK_TABLE)
    ApplyPlanningEventFormats tblAlarm, tblHistory, tblAck
    EnsureEventHistoryToggleShapes wsHistory
    EnsureEventHistoryCommandButtons wsHistory, wsAck
    RefreshEventHistoryToggleVisuals wsHistory

CleanExit:
    If internalWriteStarted Then
        gEventHistoryInternalWriteDepth = gEventHistoryInternalWriteDepth - 1
        If gEventHistoryInternalWriteDepth < 0 Then gEventHistoryInternalWriteDepth = 0
    End If
    Exit Sub

CleanFail:
    Resume CleanExit

End Sub

'------------------------------------------------------------------------------
' FR: Journalise Planning Event dans l'historique planning.
' EN: Logs Planning Event into the planning history.
'------------------------------------------------------------------------------
Public Sub LogPlanningEvent( _
    ByVal severity As String, _
    ByVal eventType As String, _
    ByVal eventHash As String, _
    ByVal frMessage As String, _
    ByVal enMessage As String, _
    ByVal frDetails As String, _
    ByVal enDetails As String, _
    ByVal sourceProcedure As String, _
    Optional ByVal sourceSheet As String = "", _
    Optional ByVal sourceTable As String = "", _
    Optional ByVal taskId As String = "", _
    Optional ByVal wbsValue As String = "", _
    Optional ByVal taskName As String = "", _
    Optional ByVal refreshView As Boolean = True)

    Dim wsAlarm As Worksheet
    Dim wsHistory As Worksheet
    Dim wsAck As Worksheet
    Dim tblAlarm As ListObject
    Dim tblHistory As ListObject
    Dim alarmRow As ListRow
    Dim eventTs As Date
    Dim finalHash As String
    Dim normalizedWbs As String
    Dim internalWriteStarted As Boolean

    On Error GoTo CleanFail
    If Trim$(severity) = "" Then Exit Sub
    If Trim$(eventType) = "" Then Exit Sub

    EnsurePlanningEventRunId sourceProcedure
    frMessage = TrimPlanningMessageLineEdges(frMessage)
    enMessage = TrimPlanningMessageLineEdges(enMessage)
    frDetails = TrimPlanningMessageLineEdges(frDetails)
    enDetails = TrimPlanningMessageLineEdges(enDetails)
    normalizedWbs = NormalizeWBS(wbsValue)

    BeginPlanningEventInternalWrite wsAlarm, wsHistory, wsAck
    internalWriteStarted = True
    Set tblAlarm = wsAlarm.ListObjects(CALC_ALARM_TABLE)
    Set tblHistory = wsHistory.ListObjects(EVENT_HISTORY_TABLE)

    eventTs = Now
    finalHash = Trim$(eventHash)
    If finalHash = "" Then
        finalHash = BuildPlanningEventHash( _
            severity, eventType, frMessage, enMessage, frDetails, enDetails, _
            sourceProcedure, sourceSheet, sourceTable, taskId, normalizedWbs, taskName)
    End If

    Set alarmRow = tblAlarm.ListRows.Add
    With alarmRow.Range
        .Cells(1, 1).value = BuildPlanningEventId(tblAlarm.ListRows.Count, eventTs)
        .Cells(1, 2).value = gPlanningEventRunId
        .Cells(1, 3).value = eventTs
        .Cells(1, 4).value = dateValue(eventTs)
        .Cells(1, 5).value = TimeValue(eventTs)
        .Cells(1, 6).value = UCase$(Trim$(severity))
        .Cells(1, 7).value = Trim$(eventType)
        .Cells(1, 8).value = finalHash
        .Cells(1, 9).value = False
        .Cells(1, 10).value = Trim$(sourceSheet)
        .Cells(1, 11).value = Trim$(sourceTable)
        .Cells(1, 12).value = Trim$(taskId)
        .Cells(1, 13).NumberFormat = "@"
        .Cells(1, 13).value = normalizedWbs
        .Cells(1, 14).value = Trim$(taskName)
        .Cells(1, 15).value = frMessage
        .Cells(1, 16).value = enMessage
        .Cells(1, 17).value = frDetails
        .Cells(1, 18).value = enDetails
        .Cells(1, 19).value = Trim$(sourceProcedure)
    End With

    If refreshView Then
        Refresh_EventHistory_View
    End If

CleanExit:
    If internalWriteStarted Then
        EndPlanningEventInternalWrite wsAlarm, wsHistory, wsAck
    End If
    Exit Sub

CleanFail:
    Resume CleanExit

End Sub

'------------------------------------------------------------------------------
' FR: Ecrit ou synchronise Planning Events Log Console Messages Safe dans le stockage possede par le domaine.
' EN: Writes or synchronizes Planning Events Log Console Messages Safe in the store owned by the domain.
'------------------------------------------------------------------------------

Public Function PlanningEvents_LogConsoleMessagesSafe( _
    ByVal consoleMessages As Collection, _
    ByVal sourceProcedure As String, _
    ByRef errorMessage As String) As Boolean

    Dim temporaryWorkflowStarted As Boolean

    On Error GoTo Fail

    errorMessage = vbNullString

    If consoleMessages Is Nothing Then
        PlanningEvents_LogConsoleMessagesSafe = True
        Exit Function
    End If

    temporaryWorkflowStarted = EnsurePlanningWorkflowStarted(sourceProcedure)

    If IsMacroRunActive() Then
        EnsurePlanningEventRunId sourceProcedure
    Else
        BeginPlanningEventRun sourceProcedure
    End If

    LogPlanningConsoleMessages consoleMessages, sourceProcedure
    PlanningEvents_LogConsoleMessagesSafe = True

CleanExit:
    If temporaryWorkflowStarted Then EndPlanningWorkflow
    Exit Function

Fail:
    errorMessage = BiMsg( _
        "Erreur pendant l'historisation de la console runtime" & vbCrLf & _
        "-> " & Err.Description, _
        "Error while logging the runtime console" & vbCrLf & _
        "-> " & Err.Description)
    PlanningEvents_LogConsoleMessagesSafe = False
    Resume CleanExit

End Function
'------------------------------------------------------------------------------
' FR: Journalise Planning Console Messages dans l'historique planning.
' EN: Logs Planning Console Messages into the planning history.
'------------------------------------------------------------------------------
Private Sub LogPlanningConsoleMessages( _
    ByVal consoleMessages As Collection, _
    Optional ByVal sourceProcedure As String = "CalcBridge_ShowPlanningConsole")

    Dim perfScope As clsPerfScope


    Dim item As Variant
    Dim severity As String
    Dim eventType As String
    Dim rawMessage As String
    Dim frMessage As String
    Dim enMessage As String
    Dim frDetails As String
    Dim enDetails As String
    Dim seenInConsole As Object
    Dim localKey As String
    Dim wsAlarm As Worksheet
    Dim wsHistory As Worksheet
    Dim wsAck As Worksheet
    Dim internalWriteStarted As Boolean

    Set perfScope = Profiler_BeginScope("LogPlanningConsoleMessages", "Event History")

    If consoleMessages Is Nothing Then Exit Sub
    If consoleMessages.Count = 0 Then Exit Sub

    On Error GoTo CleanFail
    BeginPlanningEventInternalWrite wsAlarm, wsHistory, wsAck
    internalWriteStarted = True
    EnsurePlanningEventRunId sourceProcedure
    Set seenInConsole = CreateObject("Scripting.Dictionary")

    For Each item In consoleMessages
        If Not ConsoleMessageHistoryHandled(item) Then
            severity = NormalizeConsoleEventSeverity(CStr(item("Type")))
            If severity <> "" Then
                rawMessage = CStr(item("Message"))
                SplitConsoleMessageForHistory rawMessage, frMessage, enMessage, frDetails, enDetails
                eventType = PlanningMessage_GetOptionalText(item, "EventType")
                If eventType = "" Then eventType = "CONSOLE_" & severity
                localKey = BuildConsoleLocalDedupKey(severity, frMessage, enMessage, frDetails, enDetails)

                If Not seenInConsole.Exists(localKey) Then
                    seenInConsole(localKey) = True

                    LogPlanningEvent _
                        severity, _
                        eventType, _
                        PlanningMessage_GetOptionalText(item, "Hash"), _
                        frMessage, _
                        enMessage, _
                        frDetails, _
                        enDetails, _
                        sourceProcedure, _
                        vbNullString, _
                        vbNullString, _
                        vbNullString, _
                        vbNullString, _
                        vbNullString, _
                        False
                End If
            End If
        End If
    Next item

    Refresh_EventHistory_View

CleanExit:
    If internalWriteStarted Then
        EndPlanningEventInternalWrite wsAlarm, wsHistory, wsAck
    End If
    Exit Sub

CleanFail:
    Resume CleanExit

End Sub

'------------------------------------------------------------------------------
' FR: Construit la valeur Planning Event Hash a partir des donnees fournies par l'appelant.
' EN: Builds the Planning Event Hash value from data supplied by the caller.
'------------------------------------------------------------------------------

Public Function BuildPlanningEventHash( _
    ByVal severity As String, _
    ByVal eventType As String, _
    ByVal frMessage As String, _
    ByVal enMessage As String, _
    ByVal frDetails As String, _
    ByVal enDetails As String, _
    ByVal sourceProcedure As String, _
    Optional ByVal sourceSheet As String = "", _
    Optional ByVal sourceTable As String = "", _
    Optional ByVal taskId As String = "", _
    Optional ByVal wbsValue As String = "", _
    Optional ByVal taskName As String = "") As String

    Dim rawText As String

    rawText = BuildPlanningEventSignature( _
        severity, eventType, frMessage, enMessage, frDetails, enDetails, _
        sourceSheet, sourceTable, taskId, wbsValue, taskName)

    BuildPlanningEventHash = HashPlanningEventSignature(rawText)

End Function

'------------------------------------------------------------------------------
' FR: Construit la valeur Planning Event Signature a partir des donnees fournies par l'appelant.
' EN: Builds the Planning Event Signature value from data supplied by the caller.
'------------------------------------------------------------------------------

Private Function BuildPlanningEventSignature( _
    ByVal severity As String, _
    ByVal eventType As String, _
    ByVal frMessage As String, _
    ByVal enMessage As String, _
    ByVal frDetails As String, _
    ByVal enDetails As String, _
    Optional ByVal sourceSheet As String = "", _
    Optional ByVal sourceTable As String = "", _
    Optional ByVal taskId As String = "", _
    Optional ByVal wbsValue As String = "", _
    Optional ByVal taskName As String = "") As String

    BuildPlanningEventSignature = _
        UCase$(Trim$(severity)) & "|" & _
        UCase$(Trim$(eventType)) & "|" & _
        UCase$(Trim$(sourceSheet)) & "|" & _
        UCase$(Trim$(sourceTable)) & "|" & _
        UCase$(Trim$(taskId)) & "|" & _
        UCase$(Trim$(wbsValue)) & "|" & _
        UCase$(Trim$(taskName)) & "|" & _
        NormalizePlanningEventSignatureText(FirstPlanningEventMessageLine(frMessage)) & "|" & _
        NormalizePlanningEventSignatureText(FirstPlanningEventMessageLine(enMessage)) & "|" & _
        NormalizePlanningEventSignatureText(ExtractPlanningEventContextLine(frMessage, "IDS")) & "|" & _
        NormalizePlanningEventSignatureText(ExtractPlanningEventContextLine(enMessage, "IDS")) & "|" & _
        NormalizePlanningEventSignatureText(ExtractPlanningEventContextLine(frMessage, "ID")) & "|" & _
        NormalizePlanningEventSignatureText(ExtractPlanningEventContextLine(enMessage, "ID")) & "|" & _
        NormalizePlanningEventSignatureText(ExtractPlanningEventContextLine(frMessage, "WBS")) & "|" & _
        NormalizePlanningEventSignatureText(ExtractPlanningEventContextLine(enMessage, "WBS")) & "|" & _
        NormalizePlanningEventSignatureText(ExtractPlanningEventContextLine(frMessage, "TASK")) & "|" & _
        NormalizePlanningEventSignatureText(ExtractPlanningEventContextLine(enMessage, "TASK"))

End Function

'------------------------------------------------------------------------------
' FR: Indique si Hash Planning Event Signature est vrai pour le contexte courant.
' EN: Returns whether Hash Planning Event Signature is true for the current context.
'------------------------------------------------------------------------------
Private Function HashPlanningEventSignature(ByVal rawText As String) As String

    Dim i As Long
    Dim acc As Double
    Dim nextVal As Double

    acc = 5381#
    For i = 1 To Len(rawText)
        nextVal = (acc * 33#) + AscW(Mid$(rawText, i, 1))
        acc = nextVal - (Fix(nextVal / 2147483647#) * 2147483647#)
    Next i

    HashPlanningEventSignature = "EVH-" & Format$(CLng(acc), "0000000000")

End Function

'------------------------------------------------------------------------------
' FR: Retourne la valeur First Planning Event Message Line sans modifier les donnees d'entree.
' EN: Returns the First Planning Event Message Line value without mutating input data.
'------------------------------------------------------------------------------

Private Function FirstPlanningEventMessageLine(ByVal value As String) As String

    Dim lines() As String
    Dim i As Long
    Dim txt As String

    txt = Replace(TrimPlanningMessageLineEdges(value), vbCrLf, vbLf)
    txt = Replace(txt, vbCr, vbLf)
    lines = Split(txt, vbLf)

    For i = LBound(lines) To UBound(lines)
        If Trim$(lines(i)) <> "" And Left$(Trim$(lines(i)), 2) <> "->" Then
            FirstPlanningEventMessageLine = Trim$(lines(i))
            Exit Function
        End If
    Next i

End Function

'------------------------------------------------------------------------------
' FR: Retourne la valeur Extract Planning Event Context Line sans modifier les donnees d'entree.
' EN: Returns the Extract Planning Event Context Line value without mutating input data.
'------------------------------------------------------------------------------

Private Function ExtractPlanningEventContextLine( _
    ByVal value As String, _
    ByVal labelName As String) As String

    Dim lines() As String
    Dim i As Long
    Dim txt As String
    Dim lineText As String
    Dim labelUpper As String

    txt = Replace(TrimPlanningMessageLineEdges(value), vbCrLf, vbLf)
    txt = Replace(txt, vbCr, vbLf)
    lines = Split(txt, vbLf)
    labelUpper = UCase$(Trim$(labelName))

    For i = LBound(lines) To UBound(lines)
        lineText = Trim$(lines(i))
        If UCase$(Left$(lineText, Len(labelUpper) + 1)) = labelUpper & ":" Then
            ExtractPlanningEventContextLine = Trim$(Mid$(lineText, Len(labelUpper) + 2))
            Exit Function
        End If
        If UCase$(Left$(lineText, Len(labelUpper) + 2)) = labelUpper & " :" Then
            ExtractPlanningEventContextLine = Trim$(Mid$(lineText, Len(labelUpper) + 3))
            Exit Function
        End If
    Next i

End Function

'------------------------------------------------------------------------------
' FR: Normalise Planning Event Signature Text dans un format exploitable.
' EN: Normalizes Planning Event Signature Text into a usable format.
'------------------------------------------------------------------------------
Private Function NormalizePlanningEventSignatureText(ByVal value As String) As String

    Dim txt As String

    txt = UCase$(TrimPlanningMessageLineEdges(value))
    txt = Replace(txt, vbTab, " ")

    Do While InStr(1, txt, "  ", vbBinaryCompare) > 0
        txt = Replace(txt, "  ", " ")
    Loop

    NormalizePlanningEventSignatureText = Trim$(txt)

End Function

'------------------------------------------------------------------------------
' FR: Retourne la valeur Console Message History Handled sans modifier les donnees d'entree.
' EN: Returns the Console Message History Handled value without mutating input data.
'------------------------------------------------------------------------------

Private Function ConsoleMessageHistoryHandled(ByVal item As Variant) As Boolean

    On Error Resume Next
    ConsoleMessageHistoryHandled = CBool(item("HistoryHandled"))
    On Error GoTo 0

End Function

'------------------------------------------------------------------------------
' FR: Normalise Console Event Severity dans un format exploitable.
' EN: Normalizes Console Event Severity into a usable format.
'------------------------------------------------------------------------------
Private Function NormalizeConsoleEventSeverity(ByVal msgType As String) As String

    Select Case UCase$(Trim$(msgType))
        Case "ERROR", "STOP"
            NormalizeConsoleEventSeverity = "STOP"
        Case "WARNING"
            NormalizeConsoleEventSeverity = "WARNING"
        Case "INFO"
            NormalizeConsoleEventSeverity = "INFO"
        Case Else
            NormalizeConsoleEventSeverity = vbNullString
    End Select

End Function

'------------------------------------------------------------------------------
' FR: Transforme la valeur Split Console Message For History sans modifier la semantique du message source.
' EN: Transforms the Split Console Message For History value without changing source-message semantics.
'------------------------------------------------------------------------------

Private Sub SplitConsoleMessageForHistory( _
    ByVal rawMessage As String, _
    ByRef frMessage As String, _
    ByRef enMessage As String, _
    ByRef frDetails As String, _
    ByRef enDetails As String)

    Dim txt As String
    Dim upperTxt As String
    Dim enPos As Long
    Dim enMarkerLen As Long

    txt = TrimPlanningMessageLineEdges(CStr(rawMessage))
    upperTxt = UCase$(txt)
    frDetails = vbNullString
    enDetails = vbNullString

    If Left$(upperTxt, 3) <> "FR:" Then
        frMessage = txt
        enMessage = txt
        Exit Sub
    End If

    enPos = InStr(1, upperTxt, vbCrLf & "EN:", vbTextCompare)
    enMarkerLen = Len(vbCrLf & "EN:")
    If enPos = 0 Then
        enPos = InStr(1, upperTxt, vbLf & "EN:", vbTextCompare)
        enMarkerLen = Len(vbLf & "EN:")
    End If

    If enPos = 0 Then
        frMessage = txt
        enMessage = txt
        Exit Sub
    End If

    frMessage = TrimPlanningMessageLineEdges(Mid$(txt, 4, enPos - 4))
    enMessage = TrimPlanningMessageLineEdges(Mid$(txt, enPos + enMarkerLen))

End Sub

'------------------------------------------------------------------------------
' FR: Retourne la valeur Trim Planning Message Line Edges sans modifier les donnees d'entree.
' EN: Returns the Trim Planning Message Line Edges value without mutating input data.
'------------------------------------------------------------------------------

Private Function TrimPlanningMessageLineEdges(ByVal value As String) As String

    Dim txt As String

    txt = CStr(value)

    Do While Len(txt) > 0 And (Left$(txt, 1) = vbCr Or Left$(txt, 1) = vbLf)
        txt = Mid$(txt, 2)
    Loop

    Do While Len(txt) > 0 And (Right$(txt, 1) = vbCr Or Right$(txt, 1) = vbLf)
        txt = Left$(txt, Len(txt) - 1)
    Loop

    TrimPlanningMessageLineEdges = txt

End Function

'------------------------------------------------------------------------------
' FR: Construit la valeur Console Local Dedup Key a partir des donnees fournies par l'appelant.
' EN: Builds the Console Local Dedup Key value from data supplied by the caller.
'------------------------------------------------------------------------------

Private Function BuildConsoleLocalDedupKey( _
    ByVal severity As String, _
    ByVal frMessage As String, _
    ByVal enMessage As String, _
    ByVal frDetails As String, _
    ByVal enDetails As String) As String

    BuildConsoleLocalDedupKey = _
        UCase$(Trim$(severity)) & "|" & _
        Trim$(frMessage) & "|" & _
        Trim$(enMessage) & "|" & _
        Trim$(frDetails) & "|" & _
        Trim$(enDetails)

End Function
'------------------------------------------------------------------------------
' FR: Verifie ou cree Planning Event Run Id si necessaire.
' EN: Ensures or creates Planning Event Run Id when needed.
'------------------------------------------------------------------------------
Private Sub EnsurePlanningEventRunId(ByVal sourceProcedure As String)

    If Trim$(gPlanningEventRunId) = "" Or Not gPlanningEventRunActive Then
        gPlanningEventRunId = BuildPlanningEventRunId(sourceProcedure)
        gPlanningEventRunActive = True
    End If

End Sub

'------------------------------------------------------------------------------
' FR: Construit la valeur Planning Event Run ID a partir des donnees fournies par l'appelant.
' EN: Builds the Planning Event Run ID value from data supplied by the caller.
'------------------------------------------------------------------------------

Private Function BuildPlanningEventRunId(ByVal sourceProcedure As String) As String

    BuildPlanningEventRunId = _
        "RUN-" & Format$(Now, "yyyymmdd-hhnnss") & _
        IIf(Trim$(sourceProcedure) = "", "", "-" & Trim$(sourceProcedure))

End Function

'------------------------------------------------------------------------------
' FR: Construit la valeur Planning Event ID a partir des donnees fournies par l'appelant.
' EN: Builds the Planning Event ID value from data supplied by the caller.
'------------------------------------------------------------------------------

Private Function BuildPlanningEventId(ByVal rowCount As Long, ByVal eventTs As Date) As String

    BuildPlanningEventId = _
        "EVT-" & Format$(eventTs, "yyyymmdd-hhnnss") & "-" & Format$(rowCount, "000000")

End Function

'------------------------------------------------------------------------------
' FR: Verifie ou cree Planning Event Sheet si necessaire.
' EN: Ensures or creates Planning Event Sheet when needed.
'------------------------------------------------------------------------------
Private Function EnsurePlanningEventSheet(ByVal sheetName As String) As Worksheet

    On Error Resume Next
    Set EnsurePlanningEventSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If EnsurePlanningEventSheet Is Nothing Then
        Set EnsurePlanningEventSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        EnsurePlanningEventSheet.Name = sheetName
    End If

End Function

'------------------------------------------------------------------------------
' FR: Verifie ou cree Planning Event Table si necessaire.
' EN: Ensures or creates Planning Event Table when needed.
'------------------------------------------------------------------------------
Private Sub EnsurePlanningEventTable( _
    ByVal ws As Worksheet, _
    ByVal tableName As String, _
    ByVal headers As Variant, _
    Optional ByVal headerRow As Long = 1)

    Dim tbl As ListObject
    Dim rng As Range
    Dim i As Long

    On Error Resume Next
    Set tbl = ws.ListObjects(tableName)
    On Error GoTo 0

    If tbl Is Nothing Then
        Set rng = ws.Range(ws.Cells(headerRow, 1), ws.Cells(headerRow, UBound(headers) + 1))
        For i = LBound(headers) To UBound(headers)
            rng.Cells(1, i + 1).value = CStr(headers(i))
        Next i
        Set tbl = ws.ListObjects.Add(xlSrcRange, rng, , xlYes)
        tbl.Name = tableName
    Else
        If tableName = EVENT_HISTORY_TABLE Then MigrateEventHistoryHourHeader tbl
        EnsurePlanningEventTableHeaders tbl, headers
    End If

End Sub

'------------------------------------------------------------------------------
' FR: Traite la reference Migrate Event History Hour Header sans modifier les donnees d'entree.
' EN: Handles the Migrate Event History Hour Header reference without mutating input data.
' FR - Effet de bord : ecrit dans une table Excel detenue par le workflow.
' FR - Effet de bord : efface uniquement les donnees ou objets cibles du contrat.
' EN - Side effect: writes to an Excel table owned by the workflow.
' EN - Side effect: clears only data or objects targeted by the contract.
'------------------------------------------------------------------------------

Private Sub MigrateEventHistoryHourHeader(ByVal tbl As ListObject)

    Dim legacyColumn As ListColumn
    Dim hourColumn As ListColumn
    Dim r As Long

    If tbl Is Nothing Then Exit Sub

    On Error Resume Next
    Set legacyColumn = tbl.ListColumns("Heure")
    Set hourColumn = SchemaListColumn(tbl, VTS_TABLE_EVENT_HISTORY, VTS_COL_HOUR)
    On Error GoTo 0

    If legacyColumn Is Nothing Then Exit Sub
    If Not hourColumn Is Nothing Then
        If legacyColumn.Index = hourColumn.Index Then Exit Sub
    End If

    If hourColumn Is Nothing Then
        legacyColumn.Name = "Hour"
        Exit Sub
    End If

    If Not legacyColumn.DataBodyRange Is Nothing And Not hourColumn.DataBodyRange Is Nothing Then
        For r = 1 To tbl.ListRows.Count
            If Trim$(CStr(hourColumn.DataBodyRange.Cells(r, 1).value)) = "" Then
                hourColumn.DataBodyRange.Cells(r, 1).value = legacyColumn.DataBodyRange.Cells(r, 1).value
            End If
        Next r
    End If

    legacyColumn.Delete

End Sub

'------------------------------------------------------------------------------
' FR: Verifie ou cree Event History State si necessaire.
' EN: Ensures or creates Event History State when needed.
'------------------------------------------------------------------------------
Private Sub EnsureEventHistoryState()

    If Not gEventHistoryShowInfoInitialized Then
        gEventHistoryShowInfo = Settings_GetInfoEnabled()
        gEventHistoryShowInfoInitialized = True
    End If

    If Trim$(gPlanningLanguage) = "" Then _
        gPlanningLanguage = Settings_GetOwnerLanguage("EVENT")

End Sub

'------------------------------------------------------------------------------
' FR: Construit la valeur Event History Display Message a partir des donnees fournies par l'appelant.
' EN: Builds the Event History Display Message value from data supplied by the caller.
'------------------------------------------------------------------------------

Private Function BuildEventHistoryDisplayMessage( _
    ByVal frMessage As String, _
    ByVal enMessage As String, _
    ByVal frDetails As String, _
    ByVal enDetails As String) As String

    Dim msg As String
    Dim details As String

    EnsureEventHistoryState

    If UCase$(Trim$(gPlanningLanguage)) = "EN" Then
        msg = TrimPlanningMessageLineEdges(enMessage)
        details = TrimPlanningMessageLineEdges(enDetails)
    Else
        msg = TrimPlanningMessageLineEdges(frMessage)
        details = TrimPlanningMessageLineEdges(frDetails)
    End If

    If details <> "" Then
        If msg <> "" Then
            msg = msg & vbLf & "-> " & details
        Else
            msg = details
        End If
    End If

    BuildEventHistoryDisplayMessage = msg

End Function

'------------------------------------------------------------------------------
' FR: Construit la map Event Ack Lookup a partir des donnees fournies par l'appelant.
' EN: Builds the Event Ack Lookup map from data supplied by the caller.
'------------------------------------------------------------------------------

Private Function BuildEventAckLookup(ByVal tblAck As ListObject) As Object

    Dim lookup As Object
    Dim r As Long
    Dim severity As String
    Dim eventType As String
    Dim eventHash As String

    Set lookup = CreateObject("Scripting.Dictionary")

    If tblAck Is Nothing Then
        Set BuildEventAckLookup = lookup
        Exit Function
    End If

    If tblAck.DataBodyRange Is Nothing Then
        Set BuildEventAckLookup = lookup
        Exit Function
    End If

    For r = 1 To tblAck.ListRows.Count
        severity = UCase$(Trim$(CStr(tblAck.DataBodyRange.Cells(r, SchemaListColumn(tblAck, VTS_TABLE_EVENT_ACK, VTS_COL_SEVERITY).Index).value)))
        eventType = Trim$(CStr(tblAck.DataBodyRange.Cells(r, SchemaListColumn(tblAck, VTS_TABLE_EVENT_ACK, VTS_COL_EVENT_TYPE).Index).value))
        eventHash = Trim$(CStr(tblAck.DataBodyRange.Cells(r, SchemaListColumn(tblAck, VTS_TABLE_EVENT_ACK, VTS_COL_HASH).Index).value))

        If severity = "WARNING" And IsTruthy(tblAck.DataBodyRange.Cells(r, SchemaListColumn(tblAck, VTS_TABLE_EVENT_ACK, VTS_COL_ACKNOWLEDGED).Index).value) Then
            If eventType <> "" And eventHash <> "" Then
                lookup(BuildEventAckKey(severity, eventType, eventHash)) = True
            End If
        End If
    Next r

    Set BuildEventAckLookup = lookup

End Function

'------------------------------------------------------------------------------
' FR: Indique si Event Acknowledged est vrai pour le contexte courant.
' EN: Returns whether Event Acknowledged is true for the current context.
'------------------------------------------------------------------------------
Private Function IsEventAcknowledged( _
    ByVal ackLookup As Object, _
    ByVal severity As String, _
    ByVal eventType As String, _
    ByVal eventHash As String) As Boolean

    If ackLookup Is Nothing Then Exit Function
    If UCase$(Trim$(severity)) <> "WARNING" Then Exit Function
    If Trim$(eventType) = "" Then Exit Function
    If Trim$(eventHash) = "" Then Exit Function

    IsEventAcknowledged = ackLookup.Exists(BuildEventAckKey(severity, eventType, eventHash))

End Function

'------------------------------------------------------------------------------
' FR: Construit la valeur Event Ack Key a partir des donnees fournies par l'appelant.
' EN: Builds the Event Ack Key value from data supplied by the caller.
'------------------------------------------------------------------------------

Private Function BuildEventAckKey( _
    ByVal severity As String, _
    ByVal eventType As String, _
    ByVal eventHash As String) As String

    BuildEventAckKey = _
        UCase$(Trim$(severity)) & "|" & _
        UCase$(Trim$(eventType)) & "|" & _
        Trim$(eventHash)

End Function

'------------------------------------------------------------------------------
' FR: Retourne la map Are Planning Warning Ack Tokens Acknowledged sans modifier les donnees d'entree.
' EN: Returns the Are Planning Warning Ack Tokens Acknowledged map without mutating input data.
'------------------------------------------------------------------------------

Private Function ArePlanningWarningAckTokensAcknowledged( _
    ByVal ackLookup As Object, _
    ByVal ackTokens As String) As Boolean

    Dim tokens() As String
    Dim oneToken As Variant
    Dim parts() As String
    Dim eventType As String
    Dim eventHash As String
    Dim foundToken As Boolean

    If ackLookup Is Nothing Then Exit Function

    tokens = Split(ackTokens, ";")

    For Each oneToken In tokens
        If Trim$(CStr(oneToken)) <> "" Then
            parts = Split(CStr(oneToken), "|")
            If UBound(parts) <> 1 Then Exit Function

            eventType = Trim$(parts(0))
            eventHash = Trim$(parts(1))
            If eventType = "" Or eventHash = "" Then Exit Function

            foundToken = True
            If Not ackLookup.Exists(BuildEventAckKey("WARNING", eventType, eventHash)) Then
                ArePlanningWarningAckTokensAcknowledged = False
                Exit Function
            End If
        End If
    Next oneToken

    ArePlanningWarningAckTokensAcknowledged = foundToken

End Function

'------------------------------------------------------------------------------
' FR: Traite la reference Upsert Planning Warning Ack Token sans modifier les donnees d'entree.
' EN: Handles the Upsert Planning Warning Ack Token reference without mutating input data.
' FR - Effet de bord : ecrit dans une table Excel detenue par le workflow.
' EN - Side effect: writes to an Excel table owned by the workflow.
'------------------------------------------------------------------------------

Private Sub UpsertPlanningWarningAckToken( _
    ByVal ackToken As String, _
    ByVal item As Variant)

    Dim tblAck As ListObject
    Dim parts() As String
    Dim eventType As String
    Dim eventHash As String
    Dim r As Long
    Dim targetRow As ListRow
    Dim rawMessage As String
    Dim frMessage As String
    Dim enMessage As String
    Dim frDetails As String
    Dim enDetails As String
    Dim wbsValue As String
    Dim taskName As String

    parts = Split(CStr(ackToken), "|")
    If UBound(parts) <> 1 Then Exit Sub

    eventType = Trim$(parts(0))
    eventHash = Trim$(parts(1))
    If eventType = "" Or eventHash = "" Then Exit Sub

    rawMessage = PlanningMessage_GetOptionalText(item, "Message")
    SplitConsoleMessageForHistory rawMessage, frMessage, enMessage, frDetails, enDetails
    wbsValue = NormalizeWBS(ExtractPlanningEventContextLine(frMessage, "WBS"))
    taskName = ExtractPlanningEventContextLine(frMessage, "Task")
    If taskName = "" Then taskName = ExtractPlanningEventContextLine(enMessage, "Task")

    Set tblAck = ThisWorkbook.Worksheets(EVENT_ACK_SHEET).ListObjects(EVENT_ACK_TABLE)
    If taskName = "" Then taskName = FindEventTaskNameFromAlarm(eventType, eventHash)
    If wbsValue = "" Then wbsValue = NormalizeWBS(FindEventWbsFromAlarm(eventType, eventHash))

    If Not tblAck.DataBodyRange Is Nothing Then
        For r = 1 To tblAck.ListRows.Count
            If UCase$(Trim$(CStr(tblAck.DataBodyRange.Cells(r, SchemaListColumn(tblAck, VTS_TABLE_EVENT_ACK, VTS_COL_SEVERITY).Index).value))) = "WARNING" And _
               UCase$(Trim$(CStr(tblAck.DataBodyRange.Cells(r, SchemaListColumn(tblAck, VTS_TABLE_EVENT_ACK, VTS_COL_EVENT_TYPE).Index).value))) = UCase$(eventType) And _
               Trim$(CStr(tblAck.DataBodyRange.Cells(r, SchemaListColumn(tblAck, VTS_TABLE_EVENT_ACK, VTS_COL_HASH).Index).value)) = eventHash Then

                Set targetRow = tblAck.ListRows(r)
                Exit For
            End If
        Next r
    End If

    If targetRow Is Nothing Then Set targetRow = tblAck.ListRows.Add

    With targetRow.Range
        .Cells(1, SchemaListColumn(tblAck, VTS_TABLE_EVENT_ACK, VTS_COL_HASH).Index).value = eventHash
        .Cells(1, SchemaListColumn(tblAck, VTS_TABLE_EVENT_ACK, VTS_COL_SEVERITY).Index).value = "WARNING"
        .Cells(1, SchemaListColumn(tblAck, VTS_TABLE_EVENT_ACK, VTS_COL_EVENT_TYPE).Index).value = eventType
        .Cells(1, SchemaListColumn(tblAck, VTS_TABLE_EVENT_ACK, VTS_COL_ACKNOWLEDGED).Index).value = True
        .Cells(1, SchemaListColumn(tblAck, VTS_TABLE_EVENT_ACK, VTS_COL_ACKNOWLEDGED_AT).Index).value = Now
        .Cells(1, SchemaListColumn(tblAck, VTS_TABLE_EVENT_ACK, VTS_COL_ACKNOWLEDGED_BY).Index).value = Environ$("USERNAME")
        .Cells(1, SchemaListColumn(tblAck, VTS_TABLE_EVENT_ACK, VTS_COL_WBS).Index).value = wbsValue
        .Cells(1, SchemaListColumn(tblAck, VTS_TABLE_EVENT_ACK, VTS_COL_TASK_NAME).Index).value = taskName
        .Cells(1, SchemaListColumn(tblAck, VTS_TABLE_EVENT_ACK, VTS_COL_FR_MESSAGE).Index).value = frMessage
        .Cells(1, SchemaListColumn(tblAck, VTS_TABLE_EVENT_ACK, VTS_COL_EN_MESSAGE).Index).value = enMessage
    End With

End Sub

'------------------------------------------------------------------------------
' FR: Retourne la reference Event Task Name From Alarm sans exposer de mutateur sur l'etat source.
' EN: Returns the Event Task Name From Alarm reference without exposing a mutator for source state.
'------------------------------------------------------------------------------

Private Function FindEventTaskNameFromAlarm( _
    ByVal eventType As String, _
    ByVal eventHash As String) As String

    Dim tblAlarm As ListObject
    Dim r As Long

    On Error GoTo SafeExit

    Set tblAlarm = ThisWorkbook.Worksheets(CALC_ALARM_SHEET).ListObjects(CALC_ALARM_TABLE)
    If tblAlarm.DataBodyRange Is Nothing Then Exit Function

    For r = tblAlarm.ListRows.Count To 1 Step -1
        If UCase$(Trim$(CStr(tblAlarm.DataBodyRange.Cells(r, tblAlarm.ListColumns("Severity").Index).value))) = "WARNING" And _
           UCase$(Trim$(CStr(tblAlarm.DataBodyRange.Cells(r, tblAlarm.ListColumns("Event Type").Index).value))) = UCase$(Trim$(eventType)) And _
           Trim$(CStr(tblAlarm.DataBodyRange.Cells(r, tblAlarm.ListColumns("Hash").Index).value)) = Trim$(eventHash) Then

            FindEventTaskNameFromAlarm = Trim$(CStr(tblAlarm.DataBodyRange.Cells(r, tblAlarm.ListColumns("Task Name").Index).value))
            Exit Function
        End If
    Next r

SafeExit:
End Function

'------------------------------------------------------------------------------
' FR: Retourne la reference Event WBS From Alarm sans exposer de mutateur sur l'etat source.
' EN: Returns the Event WBS From Alarm reference without exposing a mutator for source state.
'------------------------------------------------------------------------------

Private Function FindEventWbsFromAlarm( _
    ByVal eventType As String, _
    ByVal eventHash As String) As String

    Dim tblAlarm As ListObject
    Dim r As Long

    On Error GoTo SafeExit

    Set tblAlarm = ThisWorkbook.Worksheets(CALC_ALARM_SHEET).ListObjects(CALC_ALARM_TABLE)
    If tblAlarm.DataBodyRange Is Nothing Then Exit Function

    For r = tblAlarm.ListRows.Count To 1 Step -1
        If UCase$(Trim$(CStr(tblAlarm.DataBodyRange.Cells(r, tblAlarm.ListColumns("Severity").Index).value))) = "WARNING" And _
           UCase$(Trim$(CStr(tblAlarm.DataBodyRange.Cells(r, tblAlarm.ListColumns("Event Type").Index).value))) = UCase$(Trim$(eventType)) And _
           Trim$(CStr(tblAlarm.DataBodyRange.Cells(r, tblAlarm.ListColumns("Hash").Index).value)) = Trim$(eventHash) Then

            FindEventWbsFromAlarm = Trim$(CStr(tblAlarm.DataBodyRange.Cells(r, tblAlarm.ListColumns("WBS").Index).value))
            Exit Function
        End If
    Next r

SafeExit:
End Function

'------------------------------------------------------------------------------
' FR: Supprime Planning Warning Ack Token du contexte event history and acknowledgements.
' EN: Removes Planning Warning Ack Token from the event history and acknowledgements context.
'------------------------------------------------------------------------------
Private Sub RemovePlanningWarningAckToken(ByVal ackToken As String)

    Dim tblAck As ListObject
    Dim parts() As String
    Dim eventType As String
    Dim eventHash As String
    Dim r As Long

    parts = Split(CStr(ackToken), "|")
    If UBound(parts) <> 1 Then Exit Sub

    eventType = Trim$(parts(0))
    eventHash = Trim$(parts(1))
    If eventType = "" Or eventHash = "" Then Exit Sub

    Set tblAck = ThisWorkbook.Worksheets(EVENT_ACK_SHEET).ListObjects(EVENT_ACK_TABLE)
    If tblAck.DataBodyRange Is Nothing Then Exit Sub

    For r = tblAck.ListRows.Count To 1 Step -1
        If UCase$(Trim$(CStr(tblAck.DataBodyRange.Cells(r, SchemaListColumn(tblAck, VTS_TABLE_EVENT_ACK, VTS_COL_SEVERITY).Index).value))) = "WARNING" And _
           UCase$(Trim$(CStr(tblAck.DataBodyRange.Cells(r, SchemaListColumn(tblAck, VTS_TABLE_EVENT_ACK, VTS_COL_EVENT_TYPE).Index).value))) = UCase$(eventType) And _
           Trim$(CStr(tblAck.DataBodyRange.Cells(r, SchemaListColumn(tblAck, VTS_TABLE_EVENT_ACK, VTS_COL_HASH).Index).value)) = eventHash Then

            tblAck.ListRows(r).Delete
        End If
    Next r

End Sub

'------------------------------------------------------------------------------
' FR: Retourne la valeur Planning Message Get Optional Text sans modifier les donnees d'entree.
' EN: Returns the Planning Message Get Optional Text value without mutating input data.
'------------------------------------------------------------------------------

Private Function PlanningMessage_GetOptionalText( _
    ByVal item As Variant, _
    ByVal keyName As String) As String

    On Error Resume Next
    PlanningMessage_GetOptionalText = Trim$(CStr(item(keyName)))
    On Error GoTo 0

End Function

'------------------------------------------------------------------------------
' FR: Indique si Truthy est vrai pour le contexte courant.
' EN: Returns whether Truthy is true for the current context.
'------------------------------------------------------------------------------
Private Function IsTruthy(ByVal value As Variant) As Boolean

    Dim txt As String

    If VarType(value) = vbBoolean Then
        IsTruthy = CBool(value)
        Exit Function
    End If

    txt = UCase$(Trim$(CStr(value)))
    IsTruthy = (txt = "TRUE" Or txt = "YES" Or txt = "OUI" Or txt = "1")

End Function

'------------------------------------------------------------------------------
' FR: Verifie ou cree Planning Event Table Headers si necessaire.
' EN: Ensures or creates Planning Event Table Headers when needed.
'------------------------------------------------------------------------------
Private Sub EnsurePlanningEventTableHeaders(ByVal tbl As ListObject, ByVal headers As Variant)

    Dim i As Long
    Dim headerName As String

    For i = LBound(headers) To UBound(headers)
        headerName = CStr(headers(i))
        If Not TableHasColumn(tbl, headerName) Then tbl.ListColumns.Add.Name = headerName
    Next i

End Sub

'------------------------------------------------------------------------------
' FR: Retourne la reference Table Has Column sans modifier les donnees d'entree.
' EN: Returns the Table Has Column reference without mutating input data.
'------------------------------------------------------------------------------

Private Function TableHasColumn(ByVal tbl As ListObject, ByVal columnName As String) As Boolean

    On Error Resume Next
    TableHasColumn = Not tbl.ListColumns(columnName) Is Nothing
    On Error GoTo 0

End Function

'------------------------------------------------------------------------------
' FR: Actualise Apply Planning Event Formats sans modifier les regles metier qui produisent les donnees.
' EN: Refreshes Apply Planning Event Formats without changing the business rules that produce the data.
'------------------------------------------------------------------------------

Private Sub ApplyPlanningEventFormats( _
    ByVal tblAlarm As ListObject, _
    ByVal tblHistory As ListObject, _
    ByVal tblAck As ListObject)

    Dim r As Long
    Dim sev As String
    Dim dateFormat As String
    Dim dateTimeFormat As String

    On Error Resume Next

    dateFormat = Settings_GetDateNumberFormat()
    dateTimeFormat = Settings_GetDateTimeNumberFormat()

    tblAlarm.ListColumns("Event ID").Range.NumberFormat = "@"
    tblAlarm.ListColumns("Run ID").Range.NumberFormat = "@"
    tblAlarm.ListColumns("Severity").Range.NumberFormat = "@"
    tblAlarm.ListColumns("Event Type").Range.NumberFormat = "@"
    tblAlarm.ListColumns("Hash").Range.NumberFormat = "@"
    tblAlarm.ListColumns("Sheet").Range.NumberFormat = "@"
    tblAlarm.ListColumns("Table").Range.NumberFormat = "@"
    tblAlarm.ListColumns("ID").Range.NumberFormat = "@"
    tblAlarm.ListColumns("WBS").Range.NumberFormat = "@"
    tblAlarm.ListColumns("Task Name").Range.NumberFormat = "@"
    tblAlarm.ListColumns("FR Message").Range.NumberFormat = "@"
    tblAlarm.ListColumns("EN Message").Range.NumberFormat = "@"
    tblAlarm.ListColumns("FR Details").Range.NumberFormat = "@"
    tblAlarm.ListColumns("EN Details").Range.NumberFormat = "@"
    tblAlarm.ListColumns("Source Procedure").Range.NumberFormat = "@"

    If Not tblAlarm.DataBodyRange Is Nothing Then
        tblAlarm.ListColumns("Timestamp").DataBodyRange.NumberFormat = "dd/mm/yyyy hh:mm:ss"
        tblAlarm.ListColumns("Date").DataBodyRange.NumberFormat = "dd/mm/yyyy"
        tblAlarm.ListColumns("Time").DataBodyRange.NumberFormat = "hh:mm:ss"
    End If

    SchemaListColumn(tblHistory, VTS_TABLE_EVENT_HISTORY, VTS_COL_SEVERITY).Range.NumberFormat = "@"
    SchemaListColumn(tblHistory, VTS_TABLE_EVENT_HISTORY, VTS_COL_MESSAGE).Range.NumberFormat = "@"
    SchemaListColumn(tblAck, VTS_TABLE_EVENT_ACK, VTS_COL_HASH).Range.NumberFormat = "@"
    SchemaListColumn(tblAck, VTS_TABLE_EVENT_ACK, VTS_COL_SEVERITY).Range.NumberFormat = "@"
    SchemaListColumn(tblAck, VTS_TABLE_EVENT_ACK, VTS_COL_EVENT_TYPE).Range.NumberFormat = "@"
    SchemaListColumn(tblAck, VTS_TABLE_EVENT_ACK, VTS_COL_ACKNOWLEDGED_BY).Range.NumberFormat = "@"
    SchemaListColumn(tblAck, VTS_TABLE_EVENT_ACK, VTS_COL_COMMENT).Range.NumberFormat = "@"
    SchemaListColumn(tblAck, VTS_TABLE_EVENT_ACK, VTS_COL_WBS).Range.NumberFormat = "@"
    SchemaListColumn(tblAck, VTS_TABLE_EVENT_ACK, VTS_COL_TASK_NAME).Range.NumberFormat = "@"
    SchemaListColumn(tblAck, VTS_TABLE_EVENT_ACK, VTS_COL_FR_MESSAGE).Range.NumberFormat = "@"
    SchemaListColumn(tblAck, VTS_TABLE_EVENT_ACK, VTS_COL_EN_MESSAGE).Range.NumberFormat = "@"

    If Not tblHistory.DataBodyRange Is Nothing Then
        SchemaListColumn(tblHistory, VTS_TABLE_EVENT_HISTORY, VTS_COL_DATE).DataBodyRange.NumberFormat = dateFormat
        SchemaListColumn(tblHistory, VTS_TABLE_EVENT_HISTORY, VTS_COL_HOUR).DataBodyRange.NumberFormat = "hh:mm:ss"
        SchemaListColumn(tblHistory, VTS_TABLE_EVENT_HISTORY, VTS_COL_MESSAGE).DataBodyRange.WrapText = True

        For r = 1 To tblHistory.ListRows.Count
            sev = UCase$(Trim$(CStr(tblHistory.DataBodyRange.Cells(r, SchemaListColumn(tblHistory, VTS_TABLE_EVENT_HISTORY, VTS_COL_SEVERITY).Index).value)))
            Select Case sev
                Case "ERROR", "STOP"
                    tblHistory.DataBodyRange.rows(r).Interior.Color = RGB(255, 235, 238)
                Case "WARNING"
                    tblHistory.DataBodyRange.rows(r).Interior.Color = RGB(255, 248, 225)
                Case "INFO"
                    tblHistory.DataBodyRange.rows(r).Interior.Color = RGB(235, 242, 250)
                Case Else
                    tblHistory.DataBodyRange.rows(r).Interior.Pattern = xlNone
            End Select
        Next r
    End If

    If Not tblAck.DataBodyRange Is Nothing Then
        SchemaListColumn(tblAck, VTS_TABLE_EVENT_ACK, VTS_COL_ACKNOWLEDGED_AT).DataBodyRange.NumberFormat = dateTimeFormat
    End If

    tblAlarm.Range.Columns.AutoFit
    tblHistory.Range.Columns.AutoFit
    tblAck.Range.Columns.AutoFit
    SchemaListColumn(tblHistory, VTS_TABLE_EVENT_HISTORY, VTS_COL_MESSAGE).Range.ColumnWidth = 72
    SchemaListColumn(tblHistory, VTS_TABLE_EVENT_HISTORY, VTS_COL_MESSAGE).Range.WrapText = True
    tblHistory.Range.VerticalAlignment = xlTop
    tblAck.Range.VerticalAlignment = xlTop
    tblHistory.Range.EntireRow.Hidden = False

    On Error GoTo 0

End Sub

'------------------------------------------------------------------------------
' FR: Verifie ou cree Event History Toggle Shapes si necessaire.
' EN: Ensures or creates Event History Toggle Shapes when needed.
'------------------------------------------------------------------------------
Private Sub EnsureEventHistoryToggleShapes(ByVal ws As Worksheet)

    On Error Resume Next
    ws.Shapes(EVENT_HISTORY_INFO_LABEL).Delete
    ws.Shapes(EVENT_HISTORY_INFO_BG).Delete
    ws.Shapes(EVENT_HISTORY_INFO_KNOB).Delete
    ws.Shapes(EVENT_HISTORY_LANG_LABEL).Delete
    ws.Shapes(EVENT_HISTORY_LANG_BG).Delete
    ws.Shapes(EVENT_HISTORY_LANG_KNOB).Delete
    On Error GoTo 0

End Sub
'------------------------------------------------------------------------------
' FR: Verifie ou cree Event History Command Buttons si necessaire.
' EN: Ensures or creates Event History Command Buttons when needed.
'------------------------------------------------------------------------------
Private Sub EnsureEventHistoryCommandButtons( _
    ByVal wsHistory As Worksheet, _
    ByVal wsAck As Worksheet)

    AddEventHistoryCommandButton _
        wsHistory, EVENT_HISTORY_CLEAR_BUTTON, _
        "Nettoyer historique / Clear History", _
        wsHistory.Range("A2").Left, wsHistory.Range("A2").Top + 2, _
        wsHistory.Range("A2:C2").Width, 18, "ClearPlanningEventHistory"

    AddEventHistoryCommandButton _
        wsAck, EVENT_ACK_CLEAR_BUTTON, _
        "Nettoyer cache / Clear list", _
        wsAck.Range("A1").Left + 4, wsAck.Range("A1").Top + 2, _
        154, 18, "ClearPlanningWarningAcknowledgements"

End Sub

'------------------------------------------------------------------------------
' FR: Ajoute Event History Command Button a la structure cible.
' EN: Adds Event History Command Button to the target structure.
'------------------------------------------------------------------------------
Private Sub AddEventHistoryCommandButton( _
    ByVal ws As Worksheet, _
    ByVal shapeName As String, _
    ByVal caption As String, _
    ByVal leftPos As Double, _
    ByVal topPos As Double, _
    ByVal widthVal As Double, _
    ByVal heightVal As Double, _
    ByVal macroName As String)

    Dim shp As Shape

    On Error Resume Next
    Set shp = ws.Shapes(shapeName)
    On Error GoTo 0

    If shp Is Nothing Then
        Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, leftPos, topPos, widthVal, heightVal)
        shp.Name = shapeName
    End If

    shp.Left = leftPos
    shp.Top = topPos
    shp.Width = widthVal
    shp.Height = heightVal
    shp.OnAction = macroName
    shp.Placement = xlMove
    shp.Adjustments.item(1) = 0.2
    shp.Fill.ForeColor.RGB = RGB(245, 245, 245)
    shp.Line.ForeColor.RGB = RGB(160, 160, 160)
    shp.Line.Weight = 0.75
    shp.Shadow.Visible = msoFalse
    shp.TextFrame2.TextRange.Text = caption
    shp.TextFrame2.VerticalAnchor = msoAnchorMiddle
    shp.TextFrame2.MarginLeft = 6
    shp.TextFrame2.MarginRight = 6
    shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    shp.TextFrame2.TextRange.Font.Name = "Segoe UI"
    shp.TextFrame2.TextRange.Font.Size = 8
    shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(55, 55, 55)

End Sub

'------------------------------------------------------------------------------
' FR: Rafraichit Event History Toggle Visuals a partir de l'etat courant.
' EN: Refreshes Event History Toggle Visuals from the current state.
'------------------------------------------------------------------------------
Private Sub RefreshEventHistoryToggleVisuals(ByVal ws As Worksheet)

    Dim shpInfoBg As Shape
    Dim shpInfoKnob As Shape
    Dim shpInfoLabel As Shape
    Dim shpLangBg As Shape
    Dim shpLangKnob As Shape
    Dim shpLangLabel As Shape
    Dim langIsEN As Boolean

    EnsureEventHistoryState

    On Error Resume Next
    Set shpInfoBg = ws.Shapes(EVENT_HISTORY_INFO_BG)
    Set shpInfoKnob = ws.Shapes(EVENT_HISTORY_INFO_KNOB)
    Set shpInfoLabel = ws.Shapes(EVENT_HISTORY_INFO_LABEL)
    Set shpLangBg = ws.Shapes(EVENT_HISTORY_LANG_BG)
    Set shpLangKnob = ws.Shapes(EVENT_HISTORY_LANG_KNOB)
    Set shpLangLabel = ws.Shapes(EVENT_HISTORY_LANG_LABEL)
    On Error GoTo 0

    If Not shpInfoBg Is Nothing Then FormatEventHistoryToggleTrack shpInfoBg, gEventHistoryShowInfo
    If Not shpInfoKnob Is Nothing Then PositionEventHistoryToggleKnob shpInfoKnob, shpInfoBg, gEventHistoryShowInfo
    If Not shpInfoLabel Is Nothing Then FormatEventHistoryToggleLabel shpInfoLabel, gEventHistoryShowInfo

    langIsEN = (UCase$(Trim$(gPlanningLanguage)) = "EN")
    If Not shpLangBg Is Nothing Then FormatEventHistoryToggleTrack shpLangBg, langIsEN
    If Not shpLangKnob Is Nothing Then PositionEventHistoryToggleKnob shpLangKnob, shpLangBg, langIsEN
    If Not shpLangLabel Is Nothing Then FormatEventHistoryToggleLabel shpLangLabel, True

End Sub

'------------------------------------------------------------------------------
' FR: Ajoute Event History Toggle Label a la structure cible.
' EN: Adds Event History Toggle Label to the target structure.
'------------------------------------------------------------------------------
Private Sub AddEventHistoryToggleLabel( _
    ByVal ws As Worksheet, _
    ByVal shapeName As String, _
    ByVal caption As String, _
    ByVal leftPos As Double, _
    ByVal topPos As Double, _
    ByVal widthVal As Double, _
    ByVal heightVal As Double, _
    ByVal macroName As String)

    Dim shp As Shape

    On Error Resume Next
    Set shp = ws.Shapes(shapeName)
    On Error GoTo 0

    If shp Is Nothing Then
        Set shp = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, leftPos, topPos, widthVal, heightVal)
        shp.Name = shapeName
    End If

    shp.Left = leftPos
    shp.Top = topPos
    shp.Width = widthVal
    shp.Height = heightVal
    shp.TextFrame2.TextRange.Text = caption
    shp.OnAction = macroName
    shp.Placement = xlMove
    FormatEventHistoryToggleLabel shp, True

End Sub

'------------------------------------------------------------------------------
' FR: Ajoute Event History Toggle Track a la structure cible.
' EN: Adds Event History Toggle Track to the target structure.
'------------------------------------------------------------------------------
Private Sub AddEventHistoryToggleTrack( _
    ByVal ws As Worksheet, _
    ByVal shapeName As String, _
    ByVal leftPos As Double, _
    ByVal topPos As Double, _
    ByVal widthVal As Double, _
    ByVal heightVal As Double, _
    ByVal macroName As String)

    Dim shp As Shape

    On Error Resume Next
    Set shp = ws.Shapes(shapeName)
    On Error GoTo 0

    If shp Is Nothing Then
        Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, leftPos, topPos, widthVal, heightVal)
        shp.Name = shapeName
    End If

    shp.Left = leftPos
    shp.Top = topPos
    shp.Width = widthVal
    shp.Height = heightVal
    shp.OnAction = macroName
    shp.Placement = xlMove
    shp.Adjustments.item(1) = 0.5

End Sub

'------------------------------------------------------------------------------
' FR: Ajoute Event History Toggle Knob a la structure cible.
' EN: Adds Event History Toggle Knob to the target structure.
'------------------------------------------------------------------------------
Private Sub AddEventHistoryToggleKnob( _
    ByVal ws As Worksheet, _
    ByVal shapeName As String, _
    ByVal leftPos As Double, _
    ByVal topPos As Double, _
    ByVal knobSize As Double, _
    ByVal macroName As String)

    Dim shp As Shape

    On Error Resume Next
    Set shp = ws.Shapes(shapeName)
    On Error GoTo 0

    If shp Is Nothing Then
        Set shp = ws.Shapes.AddShape(msoShapeOval, leftPos, topPos, knobSize, knobSize)
        shp.Name = shapeName
        shp.Left = leftPos
        shp.Top = topPos
    End If

    shp.Width = knobSize
    shp.Height = knobSize
    shp.OnAction = macroName
    shp.Placement = xlMove
    shp.Fill.ForeColor.RGB = RGB(255, 255, 255)
    shp.Line.ForeColor.RGB = RGB(150, 150, 150)
    shp.Line.Weight = 0.75
    shp.Shadow.Visible = msoFalse

End Sub

'------------------------------------------------------------------------------
' FR: Formate Event History Toggle Track pour l'affichage ou l'ecriture.
' EN: Formats Event History Toggle Track for display or writing.
'------------------------------------------------------------------------------
Private Sub FormatEventHistoryToggleTrack(ByVal shp As Shape, ByVal isOn As Boolean)

    shp.Line.Weight = 1
    shp.Shadow.Visible = msoFalse

    If isOn Then
        shp.Fill.ForeColor.RGB = RGB(68, 114, 196)
        shp.Line.ForeColor.RGB = RGB(68, 114, 196)
    Else
        shp.Fill.ForeColor.RGB = RGB(230, 230, 230)
        shp.Line.ForeColor.RGB = RGB(170, 170, 170)
    End If

End Sub

'------------------------------------------------------------------------------
' FR: Formate Event History Toggle Label pour l'affichage ou l'ecriture.
' EN: Formats Event History Toggle Label for display or writing.
'------------------------------------------------------------------------------
Private Sub FormatEventHistoryToggleLabel(ByVal shp As Shape, ByVal isOn As Boolean)

    shp.Line.Visible = msoFalse
    shp.Fill.Visible = msoFalse
    shp.TextFrame2.VerticalAnchor = msoAnchorMiddle
    shp.TextFrame2.MarginLeft = 0
    shp.TextFrame2.MarginRight = 0
    shp.TextFrame2.TextRange.Font.Name = "Segoe UI"
    shp.TextFrame2.TextRange.Font.Size = 9
    shp.TextFrame2.TextRange.Font.Bold = msoTrue

    If isOn Then
        shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(68, 114, 196)
    Else
        shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(90, 90, 90)
    End If

End Sub

'------------------------------------------------------------------------------
' FR: Traite la reference Position Event History Toggle Knob sans modifier les donnees d'entree.
' EN: Handles the Position Event History Toggle Knob reference without mutating input data.
'------------------------------------------------------------------------------

Private Sub PositionEventHistoryToggleKnob(ByVal knob As Shape, ByVal track As Shape, ByVal isOn As Boolean)

    If track Is Nothing Then Exit Sub

    knob.Top = track.Top + ((track.Height - knob.Height) / 2)
    If isOn Then
        knob.Left = track.Left + track.Width - knob.Width - 2
    Else
        knob.Left = track.Left + 2
    End If

End Sub

'------------------------------------------------------------------------------
' FR: Supprime Event History Shape If Exists du contexte event history and acknowledgements.
' EN: Removes Event History Shape If Exists from the event history and acknowledgements context.
'------------------------------------------------------------------------------
Private Sub DeleteEventHistoryShapeIfExists(ByVal ws As Worksheet, ByVal shapeName As String)

    On Error Resume Next
    ws.Shapes(shapeName).Delete
    On Error GoTo 0

End Sub

'------------------------------------------------------------------------------
' FR: Vide ou reinitialise Planning Event Table Rows.
' EN: Clears or resets Planning Event Table Rows.
'------------------------------------------------------------------------------
Private Sub ClearPlanningEventTableRows(ByVal tbl As ListObject)

    If tbl Is Nothing Then Exit Sub
    If tbl.DataBodyRange Is Nothing Then Exit Sub

    tbl.DataBodyRange.Delete

End Sub

'------------------------------------------------------------------------------
' FR: Indique si Planning Event Internal Write Active est vrai pour le contexte courant.
' EN: Returns whether Planning Event Internal Write Active is true for the current context.
'------------------------------------------------------------------------------
Public Function IsPlanningEventInternalWriteActive() As Boolean

    IsPlanningEventInternalWriteActive = (gEventHistoryInternalWriteDepth > 0)

End Function

'------------------------------------------------------------------------------
' FR: Ouvre le cycle de traitement Planning Event Internal Write.
' EN: Begins the Planning Event Internal Write processing cycle.
'------------------------------------------------------------------------------
Private Sub BeginPlanningEventInternalWrite( _
    ByRef wsAlarm As Worksheet, _
    ByRef wsHistory As Worksheet, _
    ByRef wsAck As Worksheet)

    gEventHistoryInternalWriteDepth = gEventHistoryInternalWriteDepth + 1

    If Not IsPlanningEventInfrastructureReady() Then
        EnsurePlanningEventHistoryInfrastructure
    End If

    Set wsAlarm = ThisWorkbook.Worksheets(CALC_ALARM_SHEET)
    Set wsHistory = ThisWorkbook.Worksheets(EVENT_HISTORY_SHEET)
    Set wsAck = ThisWorkbook.Worksheets(EVENT_ACK_SHEET)

End Sub

'------------------------------------------------------------------------------
' FR: Ferme le cycle de traitement Planning Event Internal Write.
' EN: Ends the Planning Event Internal Write processing cycle.
'------------------------------------------------------------------------------
Private Sub EndPlanningEventInternalWrite( _
    ByVal wsAlarm As Worksheet, _
    ByVal wsHistory As Worksheet, _
    ByVal wsAck As Worksheet)

    If gEventHistoryInternalWriteDepth <= 0 Then
        gEventHistoryInternalWriteDepth = 0
        Exit Sub
    End If

    gEventHistoryInternalWriteDepth = gEventHistoryInternalWriteDepth - 1

End Sub

'------------------------------------------------------------------------------
' FR: Indique si Planning Event Infrastructure Ready est vrai pour le contexte courant.
' EN: Returns whether Planning Event Infrastructure Ready is true for the current context.
'------------------------------------------------------------------------------
Private Function IsPlanningEventInfrastructureReady() As Boolean

    Dim wsAlarm As Worksheet
    Dim wsHistory As Worksheet
    Dim wsAck As Worksheet
    Dim tblAlarm As ListObject
    Dim tblHistory As ListObject
    Dim tblAck As ListObject

    On Error GoTo SafeExit

    Set wsAlarm = ThisWorkbook.Worksheets(CALC_ALARM_SHEET)
    Set wsHistory = ThisWorkbook.Worksheets(EVENT_HISTORY_SHEET)
    Set wsAck = ThisWorkbook.Worksheets(EVENT_ACK_SHEET)
    Set tblAlarm = wsAlarm.ListObjects(CALC_ALARM_TABLE)
    Set tblHistory = wsHistory.ListObjects(EVENT_HISTORY_TABLE)
    Set tblAck = wsAck.ListObjects(EVENT_ACK_TABLE)

    IsPlanningEventInfrastructureReady = _
        TableHasColumn(tblAlarm, "Event ID") And _
        TableHasColumn(tblAlarm, "Source Procedure") And _
        TableHasColumn(tblHistory, "Acknowledged") And _
        TableHasColumn(tblAck, "Hash") And _
        TableHasColumn(tblAck, "EN Message")

SafeExit:
End Function

'------------------------------------------------------------------------------
' FR: Retourne la valeur Calc Alarm Headers sans modifier les donnees d'entree.
' EN: Returns the Calc Alarm Headers value without mutating input data.
'------------------------------------------------------------------------------

Private Function CalcAlarmHeaders() As Variant

    CalcAlarmHeaders = Array( _
        "Event ID", _
        "Run ID", _
        "Timestamp", _
        "Date", _
        "Time", _
        "Severity", _
        "Event Type", _
        "Hash", _
        "Acknowledged", _
        "Sheet", _
        "Table", _
        "ID", _
        "WBS", _
        "Task Name", _
        "FR Message", _
        "EN Message", _
        "FR Details", _
        "EN Details", _
        "Source Procedure")

End Function

'------------------------------------------------------------------------------
' FR: Retourne la valeur Event History Headers sans modifier les donnees d'entree.
' EN: Returns the Event History Headers value without mutating input data.
'------------------------------------------------------------------------------

Private Function EventHistoryHeaders() As Variant

    EventHistoryHeaders = Array( _
        SchemaCurrentColumnTitle(VTS_TABLE_EVENT_HISTORY, VTS_COL_DATE), _
        SchemaCurrentColumnTitle(VTS_TABLE_EVENT_HISTORY, VTS_COL_HOUR), _
        SchemaCurrentColumnTitle(VTS_TABLE_EVENT_HISTORY, VTS_COL_SEVERITY), _
        SchemaCurrentColumnTitle(VTS_TABLE_EVENT_HISTORY, VTS_COL_MESSAGE), _
        SchemaCurrentColumnTitle(VTS_TABLE_EVENT_HISTORY, VTS_COL_ACKNOWLEDGED))

End Function

'------------------------------------------------------------------------------
' FR: Retourne la valeur Event Ack Headers sans modifier les donnees d'entree.
' EN: Returns the Event Ack Headers value without mutating input data.
'------------------------------------------------------------------------------

Private Function EventAckHeaders() As Variant

    EventAckHeaders = Array( _
        SchemaCurrentColumnTitle(VTS_TABLE_EVENT_ACK, VTS_COL_HASH), _
        SchemaCurrentColumnTitle(VTS_TABLE_EVENT_ACK, VTS_COL_SEVERITY), _
        SchemaCurrentColumnTitle(VTS_TABLE_EVENT_ACK, VTS_COL_EVENT_TYPE), _
        SchemaCurrentColumnTitle(VTS_TABLE_EVENT_ACK, VTS_COL_ACKNOWLEDGED), _
        SchemaCurrentColumnTitle(VTS_TABLE_EVENT_ACK, VTS_COL_ACKNOWLEDGED_AT), _
        SchemaCurrentColumnTitle(VTS_TABLE_EVENT_ACK, VTS_COL_ACKNOWLEDGED_BY), _
        SchemaCurrentColumnTitle(VTS_TABLE_EVENT_ACK, VTS_COL_COMMENT), _
        SchemaCurrentColumnTitle(VTS_TABLE_EVENT_ACK, VTS_COL_WBS), _
        SchemaCurrentColumnTitle(VTS_TABLE_EVENT_ACK, VTS_COL_TASK_NAME), _
        SchemaCurrentColumnTitle(VTS_TABLE_EVENT_ACK, VTS_COL_FR_MESSAGE), _
        SchemaCurrentColumnTitle(VTS_TABLE_EVENT_ACK, VTS_COL_EN_MESSAGE))

End Function

'------------------------------------------------------------------------------
' FR: Vide les stockages runtime Alarm, EventHistory et ACK sans recreer l'infrastructure.
' EN: Clears Alarm, EventHistory, and ACK runtime storage without rebuilding infrastructure.
'------------------------------------------------------------------------------
Public Sub EventHistory_ResetRuntimeStorage()

    Dim wsAlarm As Worksheet
    Dim wsHistory As Worksheet
    Dim wsAck As Worksheet
    Dim tblAlarm As ListObject
    Dim tblHistory As ListObject
    Dim tblAck As ListObject

    On Error Resume Next
    Set wsAlarm = ThisWorkbook.Worksheets(CALC_ALARM_SHEET)
    Set wsHistory = ThisWorkbook.Worksheets(EVENT_HISTORY_SHEET)
    Set wsAck = ThisWorkbook.Worksheets(EVENT_ACK_SHEET)
    If Not wsAlarm Is Nothing Then Set tblAlarm = wsAlarm.ListObjects(CALC_ALARM_TABLE)
    If Not wsHistory Is Nothing Then Set tblHistory = wsHistory.ListObjects(EVENT_HISTORY_TABLE)
    If Not wsAck Is Nothing Then Set tblAck = wsAck.ListObjects(EVENT_ACK_TABLE)
    On Error GoTo 0

    If Not tblAlarm Is Nothing Then ClearPlanningEventTableRows tblAlarm
    If Not tblHistory Is Nothing Then ClearPlanningEventTableRows tblHistory
    If Not tblAck Is Nothing Then ClearPlanningEventTableRows tblAck

End Sub
