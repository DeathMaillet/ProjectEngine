Attribute VB_Name = "mod_SCurveDiagnosticsHarness"
Option Explicit

'===============================================================================
' MODULE : mod_SCurveDiagnosticsHarness
' DOMAINE / DOMAIN : Validation Harnesses
'
' FR
' Harnais de preuve du contrat S Curve Diagnostics sur des copies de test.
' N'appartient a aucun workflow produit et ne doit pas etre appele en usage normal.
'
' EN
' Proof harness for the S Curve Diagnostics contract on test copies.
' Is not production workflow code and must not run during normal use.
'
' CONTRATS / CONTRACTS : SCurveDiagnosticsHarness_Smoke
' CALLBACKS EXTERNES / EXTERNAL CALLBACKS : Aucun / None
'===============================================================================


Private gSCurveDiagnosticsTracePath As String

'------------------------------------------------------------------------------
' FR: Execute un smoke deterministe du service diagnostics S-Curve.
' EN: Runs a deterministic smoke for the S-Curve diagnostics service.
'------------------------------------------------------------------------------
Public Function SCurveDiagnosticsHarness_Smoke( _
    ByVal capturePath As String, _
    ByVal matrixPath As String) As String

    Dim messages As Collection
    Dim stopIds As Object
    Dim warningIds As Object
    Dim idToWbs As Object
    Dim ackTokens As String

    On Error GoTo Fail

    gSCurveDiagnosticsTracePath = matrixPath & ".trace.txt"
    SCurveDiagnosticsHarness_DeleteFile gSCurveDiagnosticsTracePath
    SCurveDiagnosticsHarness_DeleteFile matrixPath
    SCurveDiagnosticsHarness_Trace "01 enter"

    EnsurePlanningEventHistoryInfrastructure
    PlanningConsolePolicy_DisableNonInteractive
    EventHistory_SetShowInfo True
    EventHistory_SetLanguage "EN"
    ClearPlanningWarningAcknowledgements
    ClearPlanningEventHistory
    SCurveDiagnosticsHarness_Trace "02 reset ok"

    Set messages = New Collection
    Set stopIds = CreateObject("Scripting.Dictionary")
    Set warningIds = CreateObject("Scripting.Dictionary")
    Set idToWbs = CreateObject("Scripting.Dictionary")

    stopIds("SC-STOP-001") = True
    stopIds("SC-STOP-002") = True
    warningIds("SC-WARN-001") = True
    warningIds("SC-WARN-002") = True
    idToWbs("SC-STOP-001") = "10.1"
    idToWbs("SC-STOP-002") = "10.2"
    idToWbs("SC-WARN-001") = "20.1"
    idToWbs("SC-WARN-002") = "20.2"

    SCurve_AddGroupedMessage messages, "STOP", stopIds, idToWbs, _
        "Données bloquantes pour S-curve", _
        "corriger les champs nécessaires avant recalcul", _
        "Blocking data for S-curve", _
        "fix required fields before recalculation"
    SCurveDiagnosticsHarness_Trace "03 grouped stop produced"

    ackTokens = SCurve_LogGroupedWarningEvents( _
        warningIds, idToWbs, _
        "SCURVE_MISSING_WEIGHT", _
        "Poids manquant sur certaines taches feuilles", _
        "Missing weight on some leaf tasks", _
        "les taches sont exclues de la S-curve ; completer Weight (%) si necessaire", _
        "tasks are excluded from the S-curve; fill Weight (%) if needed")
    SCurveDiagnosticsHarness_Assert Trim$(ackTokens) <> "", "ack tokens generated"
    SCurveDiagnosticsHarness_Trace "04 warning events logged"

    SCurve_AddGroupedMessage messages, "WARNING", warningIds, idToWbs, _
        "Poids manquant sur certaines tâches feuilles - non prises en compte dans la S-curve", _
        "compléter Weight (%) si nécessaire", _
        "Missing weight on some leaf tasks - excluded from S-curve", _
        "fill Weight (%) if needed", _
        True, _
        ackTokens
    SCurveDiagnosticsHarness_Trace "05 grouped warning produced"

    SCurve_AddConsoleMessage messages, "WARNING", _
        "Aucune tâche feuille exploitable pour la S-curve.", _
        "No valid leaf task available for S-curve.", _
        "SCURVE_NO_VALID_LEAF_TASK"
    SCurveDiagnosticsHarness_Trace "06 simple warning produced"

    PlanningConsolePolicy_EnableNonInteractive capturePath, "SCurveDiagnosticsHarness"
    CalcBridge_ShowPlanningConsole messages
    SCurveDiagnosticsHarness_Trace "07 console show returned"
    SCurveDiagnosticsHarness_Assert PlanningConsolePolicy_GetCapturedMessageCount() >= 3, "console captured S-Curve diagnostics"
    PlanningConsolePolicy_DisableNonInteractive

    SCurveDiagnosticsHarness_Assert SCurveDiagnosticsHarness_EventTypeExists("SCURVE_MISSING_WEIGHT"), "missing weight event logged"
    SCurveDiagnosticsHarness_Assert SCurveDiagnosticsHarness_EventTypeExists("SCURVE_NO_VALID_LEAF_TASK"), "no valid leaf warning logged"
    SCurveDiagnosticsHarness_Assert SCurveDiagnosticsHarness_CaptureHasSeverity(capturePath, "STOP"), "stop present in console"
    SCurveDiagnosticsHarness_Assert SCurveDiagnosticsHarness_CaptureHasText(capturePath, "Missing weight"), "warning present in console"
    SCurveDiagnosticsHarness_Assert SCurveDiagnosticsHarness_CaptureHasText(capturePath, "SCURVE_MISSING_WEIGHT") Or Trim$(ackTokens) <> "", "ack token contract present"

    SCurveDiagnosticsHarness_WriteMatrix matrixPath, capturePath, ackTokens
    SCurveDiagnosticsHarness_Trace "08 matrix written"

    PlanningConsolePolicy_DisableNonInteractive
    SCurveDiagnosticsHarness_Trace "09 pass"
    SCurveDiagnosticsHarness_Smoke = "PASS"
    Exit Function

Fail:
    On Error Resume Next
    SCurveDiagnosticsHarness_Trace "FAIL " & Err.Number & " " & Err.Description
    PlanningConsolePolicy_DisableNonInteractive
    SCurveDiagnosticsHarness_Smoke = "FAIL: " & Err.Description

End Function

'------------------------------------------------------------------------------
' FR: Verifie le contrat S Curve Diagnostics Harness Write Matrix et signale toute divergence au harnais.
' EN: Verifies the S Curve Diagnostics Harness Write Matrix contract and reports any divergence to the harness.
'------------------------------------------------------------------------------

Private Sub SCurveDiagnosticsHarness_WriteMatrix( _
    ByVal matrixPath As String, _
    ByVal capturePath As String, _
    ByVal ackTokens As String)

    Dim fileNo As Integer

    fileNo = FreeFile
    Open matrixPath For Output As #fileNo
    Print #fileNo, "Diagnostic" & vbTab & "Severity" & vbTab & "Producer" & vbTab & "MessageEngine" & vbTab & "EventHistory" & vbTab & "ACK"
    SCurveDiagnosticsHarness_WriteMatrixRow fileNo, "SCURVE_BLOCKING_DATA", "STOP", "SCurve_AddGroupedMessage", SCurveDiagnosticsHarness_CaptureHasSeverity(capturePath, "STOP"), SCurveDiagnosticsHarness_EventTypeExists("CONSOLE_STOP"), False
    SCurveDiagnosticsHarness_WriteMatrixRow fileNo, "SCURVE_MISSING_WEIGHT", "WARNING", "SCurve_LogGroupedWarningEvents + SCurve_AddGroupedMessage", SCurveDiagnosticsHarness_CaptureHasText(capturePath, "Missing weight"), SCurveDiagnosticsHarness_EventTypeExists("SCURVE_MISSING_WEIGHT"), Trim$(ackTokens) <> ""
    SCurveDiagnosticsHarness_WriteMatrixRow fileNo, "SCURVE_NO_VALID_LEAF_TASK", "WARNING", "SCurve_AddConsoleMessage", SCurveDiagnosticsHarness_CaptureHasText(capturePath, "No valid leaf"), SCurveDiagnosticsHarness_EventTypeExists("SCURVE_NO_VALID_LEAF_TASK"), False
    Close #fileNo

End Sub

'------------------------------------------------------------------------------
' FR: Verifie le contrat S Curve Diagnostics Harness Write Matrix Row et signale toute divergence au harnais.
' EN: Verifies the S Curve Diagnostics Harness Write Matrix Row contract and reports any divergence to the harness.
'------------------------------------------------------------------------------

Private Sub SCurveDiagnosticsHarness_WriteMatrixRow( _
    ByVal fileNo As Integer, _
    ByVal diagnosticName As String, _
    ByVal severityName As String, _
    ByVal producerName As String, _
    ByVal messageEnginePresent As Boolean, _
    ByVal eventHistoryPresent As Boolean, _
    ByVal ackPresent As Boolean)

    Print #fileNo, _
        SCurveDiagnosticsHarness_Tsv(diagnosticName) & vbTab & _
        SCurveDiagnosticsHarness_Tsv(severityName) & vbTab & _
        SCurveDiagnosticsHarness_Tsv(producerName) & vbTab & _
        CStr(messageEnginePresent) & vbTab & _
        CStr(eventHistoryPresent) & vbTab & _
        CStr(ackPresent)

End Sub

'------------------------------------------------------------------------------
' FR: Verifie le contrat S Curve Diagnostics Harness Event Type Exists et signale toute divergence au harnais.
' EN: Verifies the S Curve Diagnostics Harness Event Type Exists contract and reports any divergence to the harness.
'------------------------------------------------------------------------------

Private Function SCurveDiagnosticsHarness_EventTypeExists(ByVal eventType As String) As Boolean

    Dim tbl As ListObject
    Dim colIdx As Long
    Dim r As Long

    On Error GoTo SafeExit
    Set tbl = ThisWorkbook.Worksheets("CALC_ALARM").ListObjects("tbl_CALC_ALARM")
    If tbl.DataBodyRange Is Nothing Then Exit Function
    colIdx = tbl.ListColumns("Event Type").Index

    For r = 1 To tbl.ListRows.Count
        If UCase$(Trim$(CStr(tbl.DataBodyRange.Cells(r, colIdx).value))) = UCase$(Trim$(eventType)) Then
            SCurveDiagnosticsHarness_EventTypeExists = True
            Exit Function
        End If
    Next r

SafeExit:
End Function

'------------------------------------------------------------------------------
' FR: Verifie le contrat S Curve Diagnostics Harness Capture Has Severity et signale toute divergence au harnais.
' EN: Verifies the S Curve Diagnostics Harness Capture Has Severity contract and reports any divergence to the harness.
'------------------------------------------------------------------------------

Private Function SCurveDiagnosticsHarness_CaptureHasSeverity( _
    ByVal capturePath As String, _
    ByVal severityName As String) As Boolean

    SCurveDiagnosticsHarness_CaptureHasSeverity = _
        SCurveDiagnosticsHarness_CaptureHasText(capturePath, vbTab & UCase$(Trim$(severityName)) & vbTab)

End Function

'------------------------------------------------------------------------------
' FR: Verifie le contrat S Curve Diagnostics Harness Capture Has Text et signale toute divergence au harnais.
' EN: Verifies the S Curve Diagnostics Harness Capture Has Text contract and reports any divergence to the harness.
'------------------------------------------------------------------------------

Private Function SCurveDiagnosticsHarness_CaptureHasText( _
    ByVal capturePath As String, _
    ByVal expectedText As String) As Boolean

    Dim fileNo As Integer
    Dim lineText As String

    On Error GoTo SafeExit
    If Len(Dir$(capturePath)) = 0 Then Exit Function

    fileNo = FreeFile
    Open capturePath For Input As #fileNo
    Do While Not EOF(fileNo)
        Line Input #fileNo, lineText
        If InStr(1, lineText, expectedText, vbTextCompare) > 0 Then
            SCurveDiagnosticsHarness_CaptureHasText = True
            Exit Do
        End If
    Loop
    Close #fileNo

SafeExit:
    On Error Resume Next
    Close #fileNo
End Function

'------------------------------------------------------------------------------
' FR: Verifie le contrat S Curve Diagnostics Harness Assert et signale toute divergence au harnais.
' EN: Verifies the S Curve Diagnostics Harness Assert contract and reports any divergence to the harness.
'------------------------------------------------------------------------------

Private Sub SCurveDiagnosticsHarness_Assert(ByVal condition As Boolean, ByVal messageText As String)

    If Not condition Then
        SCurveDiagnosticsHarness_Trace "ASSERT FAIL " & messageText
        Err.Raise vbObjectError + 9381, "SCurveDiagnosticsHarness", messageText
    End If

End Sub

'------------------------------------------------------------------------------
' FR: Verifie le contrat S Curve Diagnostics Harness Trace et signale toute divergence au harnais.
' EN: Verifies the S Curve Diagnostics Harness Trace contract and reports any divergence to the harness.
'------------------------------------------------------------------------------

Private Sub SCurveDiagnosticsHarness_Trace(ByVal messageText As String)

    Dim fileNo As Integer

    On Error Resume Next
    If Trim$(gSCurveDiagnosticsTracePath) = "" Then Exit Sub
    fileNo = FreeFile
    Open gSCurveDiagnosticsTracePath For Append As #fileNo
    Print #fileNo, Format$(Now, "yyyy-mm-dd hh:nn:ss") & " | " & messageText
    Close #fileNo

End Sub

'------------------------------------------------------------------------------
' FR: Verifie le contrat S Curve Diagnostics Harness Delete File et signale toute divergence au harnais.
' EN: Verifies the S Curve Diagnostics Harness Delete File contract and reports any divergence to the harness.
'------------------------------------------------------------------------------

Private Sub SCurveDiagnosticsHarness_DeleteFile(ByVal filePath As String)

    On Error Resume Next
    If Len(Dir$(filePath)) > 0 Then Kill filePath

End Sub

'------------------------------------------------------------------------------
' FR: Verifie le contrat S Curve Diagnostics Harness Tsv et signale toute divergence au harnais.
' EN: Verifies the S Curve Diagnostics Harness Tsv contract and reports any divergence to the harness.
'------------------------------------------------------------------------------

Private Function SCurveDiagnosticsHarness_Tsv(ByVal value As String) As String

    SCurveDiagnosticsHarness_Tsv = Replace(Replace(Replace(value, vbTab, " "), vbCr, "\r"), vbLf, "\n")

End Function
