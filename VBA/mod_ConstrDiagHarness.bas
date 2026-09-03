Attribute VB_Name = "mod_ConstrDiagHarness"
Option Explicit

'===============================================================================
' MODULE : mod_ConstrDiagHarness
' DOMAINE / DOMAIN : Validation Harnesses
'
' FR
' Harnais de preuve du contrat Constr Diag sur des copies de test.
' N'appartient a aucun workflow produit et ne doit pas etre appele en usage normal.
'
' EN
' Proof harness for the Constr Diag contract on test copies.
' Is not production workflow code and must not run during normal use.
'
' CONTRATS / CONTRACTS : ConstraintsDiagnosticsHarness_Smoke
' CALLBACKS EXTERNES / EXTERNAL CALLBACKS : Aucun / None
'===============================================================================


Private gConstraintsDiagnosticsTracePath As String

'------------------------------------------------------------------------------
' FR: Execute un smoke deterministe des diagnostics Constraints.
' EN: Runs a deterministic smoke for Constraints diagnostics.
'------------------------------------------------------------------------------
Public Function ConstraintsDiagnosticsHarness_Smoke( _
    ByVal capturePath As String, _
    ByVal matrixPath As String) As String

    Dim messages As Collection
    Dim arrConstraints(1 To 2, 1 To 4) As Variant
    Dim mapConstraints As Object
    Dim stopMessage As String
    Dim warningMessage As String
    Dim eventHash As String

    On Error GoTo Fail

    gConstraintsDiagnosticsTracePath = matrixPath & ".trace.txt"
    ConstraintsDiagnosticsHarness_DeleteFile gConstraintsDiagnosticsTracePath
    ConstraintsDiagnosticsHarness_DeleteFile matrixPath
    ConstraintsDiagnosticsHarness_Trace "01 enter"

    EnsurePlanningEventHistoryInfrastructure
    PlanningConsolePolicy_DisableNonInteractive
    EventHistory_SetShowInfo True
    EventHistory_SetLanguage "EN"
    ClearPlanningWarningAcknowledgements
    ClearPlanningEventHistory
    ConstraintsDiagnosticsHarness_Trace "02 reset ok"

    Set messages = New Collection
    Set mapConstraints = CreateObject("Scripting.Dictionary")
    mapConstraints(VTS_COL_ID) = 1
    mapConstraints(VTS_COL_WBS) = 2
    mapConstraints(VTS_COL_TASK_NAME) = 3
    mapConstraints(VTS_COL_ACTIVE) = 4

    arrConstraints(1, 1) = "CSTOP-001"
    arrConstraints(1, 2) = "31.9.1"
    arrConstraints(1, 3) = "Harness invalid constraint"
    arrConstraints(1, 4) = "Yes"

    arrConstraints(2, 1) = "CWARN-001"
    arrConstraints(2, 2) = "31.9.2"
    arrConstraints(2, 3) = "Harness summary constraint"
    arrConstraints(2, 4) = "Yes"

    stopMessage = BuildConstraintValidationMessage( _
        arrConstraints, 1, mapConstraints, _
        "Contrainte active sur un ID absent de CALC", _
        "Active constraint references an ID not found in CALC")
    CalcBridge_AddConsoleMessage messages, "STOP", stopMessage
    ConstraintsDiagnosticsHarness_Trace "03 stop produced"

    eventHash = BuildPlanningEventHash( _
        "WARNING", _
        "CONSTRAINT_PARENT_IGNORED", _
        "Contrainte active ignoree sur une tache parent", _
        "Active constraint ignored on a summary task", _
        "les contraintes sur taches parent ne sont pas exportees vers CALC", _
        "constraints on summary tasks are not exported to CALC", _
        "ConstraintsDiagnosticsHarness", _
        "CONSTRAINTS", _
        "tbl_CONSTRAINTS", _
        CStr(arrConstraints(2, 1)), _
        CStr(arrConstraints(2, 2)), _
        CStr(arrConstraints(2, 3)))

    LogPlanningEvent _
        "WARNING", _
        "CONSTRAINT_PARENT_IGNORED", _
        eventHash, _
        "Contrainte active ignoree sur une tache parent", _
        "Active constraint ignored on a summary task", _
        "les contraintes sur taches parent ne sont pas exportees vers CALC", _
        "constraints on summary tasks are not exported to CALC", _
        "ConstraintsDiagnosticsHarness", _
        "CONSTRAINTS", _
        "tbl_CONSTRAINTS", _
        CStr(arrConstraints(2, 1)), _
        CStr(arrConstraints(2, 2)), _
        CStr(arrConstraints(2, 3)), _
        False

    warningMessage = BuildConstraintValidationMessage( _
        arrConstraints, 2, mapConstraints, _
        "Contrainte active ignoree sur une tache parent", _
        "Active constraint ignored on a summary task", _
        "les contraintes sur taches parent ne sont pas exportees vers CALC", _
        "constraints on summary tasks are not exported to CALC")
    AddConstraintWarning messages, warningMessage, True, "CONSTRAINT_PARENT_IGNORED", eventHash
    ConstraintsDiagnosticsHarness_Trace "04 warning produced"

    PlanningConsolePolicy_EnableNonInteractive capturePath, "ConstraintsDiagnosticsHarness"
    CalcBridge_ShowPlanningConsole messages
    ConstraintsDiagnosticsHarness_Trace "05 console show returned"
    ConstraintsDiagnosticsHarness_Assert PlanningConsolePolicy_GetCapturedMessageCount() >= 2, "console captured Constraints diagnostics"
    PlanningConsolePolicy_DisableNonInteractive

    ConstraintsDiagnosticsHarness_Assert ConstraintsDiagnosticsHarness_EventTypeExists("CONSOLE_STOP"), "stop logged through EventHistory"
    ConstraintsDiagnosticsHarness_Assert ConstraintsDiagnosticsHarness_EventTypeExists("CONSTRAINT_PARENT_IGNORED"), "constraint warning logged through EventHistory"
    ConstraintsDiagnosticsHarness_Assert ConstraintsDiagnosticsHarness_CaptureHasSeverity(capturePath, "STOP"), "stop present in console capture"
    ConstraintsDiagnosticsHarness_Assert ConstraintsDiagnosticsHarness_CaptureHasSeverity(capturePath, "WARNING"), "warning present in console capture"
    ConstraintsDiagnosticsHarness_Assert ConstraintsDiagnosticsHarness_CaptureHasText(capturePath, "Active constraint ignored on a summary task"), "warning text present"
    ConstraintsDiagnosticsHarness_Assert ConstraintsDiagnosticsHarness_CaptureHasText(capturePath, "CONSTRAINT_PARENT_IGNORED"), "ack/event token visible in capture"

    ConstraintsDiagnosticsHarness_WriteMatrix matrixPath, capturePath
    ConstraintsDiagnosticsHarness_Trace "06 matrix written"

    PlanningConsolePolicy_DisableNonInteractive
    ConstraintsDiagnosticsHarness_Trace "07 pass"
    ConstraintsDiagnosticsHarness_Smoke = "PASS"
    Exit Function

Fail:
    On Error Resume Next
    ConstraintsDiagnosticsHarness_Trace "FAIL " & Err.Number & " " & Err.Description
    PlanningConsolePolicy_DisableNonInteractive
    ConstraintsDiagnosticsHarness_Smoke = "FAIL: " & Err.Description

End Function

'------------------------------------------------------------------------------
' FR: Verifie le contrat Constraints Diagnostics Harness Write Matrix et signale toute divergence au harnais.
' EN: Verifies the Constraints Diagnostics Harness Write Matrix contract and reports any divergence to the harness.
'------------------------------------------------------------------------------

Private Sub ConstraintsDiagnosticsHarness_WriteMatrix( _
    ByVal matrixPath As String, _
    ByVal capturePath As String)

    Dim fileNo As Integer

    fileNo = FreeFile
    Open matrixPath For Output As #fileNo
    Print #fileNo, "Diagnostic" & vbTab & "Severity" & vbTab & "Producer" & vbTab & "MessageEngine" & vbTab & "EventHistory" & vbTab & "ACK"
    ConstraintsDiagnosticsHarness_WriteMatrixRow fileNo, "CONSTRAINT_VALIDATION_STOP", "STOP", "BuildConstraintValidationMessage + CalcBridge_AddConsoleMessage", ConstraintsDiagnosticsHarness_CaptureHasSeverity(capturePath, "STOP"), ConstraintsDiagnosticsHarness_EventTypeExists("CONSOLE_STOP"), False
    ConstraintsDiagnosticsHarness_WriteMatrixRow fileNo, "CONSTRAINT_PARENT_IGNORED", "WARNING", "LogPlanningEvent + AddConstraintWarning", ConstraintsDiagnosticsHarness_CaptureHasSeverity(capturePath, "WARNING"), ConstraintsDiagnosticsHarness_EventTypeExists("CONSTRAINT_PARENT_IGNORED"), ConstraintsDiagnosticsHarness_CaptureHasText(capturePath, "CONSTRAINT_PARENT_IGNORED")
    Close #fileNo

End Sub

'------------------------------------------------------------------------------
' FR: Verifie le contrat Constraints Diagnostics Harness Write Matrix Row et signale toute divergence au harnais.
' EN: Verifies the Constraints Diagnostics Harness Write Matrix Row contract and reports any divergence to the harness.
'------------------------------------------------------------------------------

Private Sub ConstraintsDiagnosticsHarness_WriteMatrixRow( _
    ByVal fileNo As Integer, _
    ByVal diagnosticName As String, _
    ByVal severityName As String, _
    ByVal producerName As String, _
    ByVal messageEnginePresent As Boolean, _
    ByVal eventHistoryPresent As Boolean, _
    ByVal ackPresent As Boolean)

    Print #fileNo, _
        ConstraintsDiagnosticsHarness_Tsv(diagnosticName) & vbTab & _
        ConstraintsDiagnosticsHarness_Tsv(severityName) & vbTab & _
        ConstraintsDiagnosticsHarness_Tsv(producerName) & vbTab & _
        CStr(messageEnginePresent) & vbTab & _
        CStr(eventHistoryPresent) & vbTab & _
        CStr(ackPresent)

End Sub

'------------------------------------------------------------------------------
' FR: Verifie le contrat Constraints Diagnostics Harness Event Type Exists et signale toute divergence au harnais.
' EN: Verifies the Constraints Diagnostics Harness Event Type Exists contract and reports any divergence to the harness.
'------------------------------------------------------------------------------

Private Function ConstraintsDiagnosticsHarness_EventTypeExists(ByVal eventType As String) As Boolean

    Dim tbl As ListObject
    Dim colIdx As Long
    Dim r As Long

    On Error GoTo SafeExit
    Set tbl = ThisWorkbook.Worksheets("CALC_ALARM").ListObjects("tbl_CALC_ALARM")
    If tbl.DataBodyRange Is Nothing Then Exit Function
    colIdx = tbl.ListColumns("Event Type").Index

    For r = 1 To tbl.ListRows.Count
        If UCase$(Trim$(CStr(tbl.DataBodyRange.Cells(r, colIdx).value))) = UCase$(Trim$(eventType)) Then
            ConstraintsDiagnosticsHarness_EventTypeExists = True
            Exit Function
        End If
    Next r

SafeExit:
End Function

'------------------------------------------------------------------------------
' FR: Verifie le contrat Constraints Diagnostics Harness Capture Has Severity et signale toute divergence au harnais.
' EN: Verifies the Constraints Diagnostics Harness Capture Has Severity contract and reports any divergence to the harness.
'------------------------------------------------------------------------------

Private Function ConstraintsDiagnosticsHarness_CaptureHasSeverity( _
    ByVal capturePath As String, _
    ByVal severityName As String) As Boolean

    ConstraintsDiagnosticsHarness_CaptureHasSeverity = _
        ConstraintsDiagnosticsHarness_CaptureHasText(capturePath, vbTab & UCase$(Trim$(severityName)) & vbTab)

End Function

'------------------------------------------------------------------------------
' FR: Verifie le contrat Constraints Diagnostics Harness Capture Has Text et signale toute divergence au harnais.
' EN: Verifies the Constraints Diagnostics Harness Capture Has Text contract and reports any divergence to the harness.
'------------------------------------------------------------------------------

Private Function ConstraintsDiagnosticsHarness_CaptureHasText( _
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
            ConstraintsDiagnosticsHarness_CaptureHasText = True
            Exit Do
        End If
    Loop
    Close #fileNo

SafeExit:
    On Error Resume Next
    Close #fileNo
End Function

'------------------------------------------------------------------------------
' FR: Verifie le contrat Constraints Diagnostics Harness Assert et signale toute divergence au harnais.
' EN: Verifies the Constraints Diagnostics Harness Assert contract and reports any divergence to the harness.
'------------------------------------------------------------------------------

Private Sub ConstraintsDiagnosticsHarness_Assert(ByVal condition As Boolean, ByVal messageText As String)

    If Not condition Then
        ConstraintsDiagnosticsHarness_Trace "ASSERT FAIL " & messageText
        Err.Raise vbObjectError + 9391, "ConstraintsDiagnosticsHarness", messageText
    End If

End Sub

'------------------------------------------------------------------------------
' FR: Verifie le contrat Constraints Diagnostics Harness Trace et signale toute divergence au harnais.
' EN: Verifies the Constraints Diagnostics Harness Trace contract and reports any divergence to the harness.
'------------------------------------------------------------------------------

Private Sub ConstraintsDiagnosticsHarness_Trace(ByVal messageText As String)

    Dim fileNo As Integer

    On Error Resume Next
    If Trim$(gConstraintsDiagnosticsTracePath) = "" Then Exit Sub
    fileNo = FreeFile
    Open gConstraintsDiagnosticsTracePath For Append As #fileNo
    Print #fileNo, Format$(Now, "yyyy-mm-dd hh:nn:ss") & " | " & messageText
    Close #fileNo

End Sub

'------------------------------------------------------------------------------
' FR: Verifie le contrat Constraints Diagnostics Harness Delete File et signale toute divergence au harnais.
' EN: Verifies the Constraints Diagnostics Harness Delete File contract and reports any divergence to the harness.
'------------------------------------------------------------------------------

Private Sub ConstraintsDiagnosticsHarness_DeleteFile(ByVal filePath As String)

    On Error Resume Next
    If Len(Dir$(filePath)) > 0 Then Kill filePath

End Sub

'------------------------------------------------------------------------------
' FR: Verifie le contrat Constraints Diagnostics Harness Tsv et signale toute divergence au harnais.
' EN: Verifies the Constraints Diagnostics Harness Tsv contract and reports any divergence to the harness.
'------------------------------------------------------------------------------

Private Function ConstraintsDiagnosticsHarness_Tsv(ByVal value As String) As String

    ConstraintsDiagnosticsHarness_Tsv = Replace(Replace(Replace(value, vbTab, " "), vbCr, "\r"), vbLf, "\n")

End Function

