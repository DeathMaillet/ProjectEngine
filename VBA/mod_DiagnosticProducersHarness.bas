Attribute VB_Name = "mod_DiagnosticProducersHarness"
Option Explicit

'===============================================================================
' MODULE : mod_DiagnosticProducersHarness
' DOMAINE / DOMAIN : Validation Harnesses
'
' FR
' Harnais de preuve du contrat Diagnostic Producers sur des copies de test.
' N'appartient a aucun workflow produit et ne doit pas etre appele en usage normal.
'
' EN
' Proof harness for the Diagnostic Producers contract on test copies.
' Is not production workflow code and must not run during normal use.
'
' CONTRATS / CONTRACTS : DiagnosticProducersHarness_Smoke
' CALLBACKS EXTERNES / EXTERNAL CALLBACKS : Aucun / None
'===============================================================================


Private gDiagnosticProducersTracePath As String

'------------------------------------------------------------------------------
' FR: Execute un smoke contractuel des producteurs de diagnostics.
' EN: Runs a contractual smoke for diagnostic producers.
'------------------------------------------------------------------------------
Public Function DiagnosticProducersHarness_Smoke( _
    ByVal capturePath As String, _
    ByVal matrixPath As String) As String

    Dim messages As Collection
    Dim eventHash As String
    Dim ackToken As String
    Dim historyBefore As Long
    Dim historyAfter As Long
    Dim coreData(1 To 4, 1 To 5) As Variant
    Dim testData(1 To 1, 1 To 5) As Variant
    Dim mapCore As Object
    Dim constraintDiagnostics As Object
    Dim dependencyDiagnostics As Object
    Dim diag As Object

    On Error GoTo Fail

    gDiagnosticProducersTracePath = matrixPath & ".trace.txt"
    DiagnosticProducersHarness_DeleteFile gDiagnosticProducersTracePath
    DiagnosticProducersHarness_DeleteFile matrixPath
    DiagnosticProducersHarness_Trace "01 enter"

    EnsurePlanningEventHistoryInfrastructure
    PlanningConsolePolicy_DisableNonInteractive
    EventHistory_SetShowInfo True
    EventHistory_SetLanguage "EN"
    ClearPlanningWarningAcknowledgements
    ClearPlanningEventHistory
    DiagnosticProducersHarness_Trace "02 reset ok"

    Set messages = New Collection

    CalcBridge_AddConsoleMessage messages, "STOP", _
        "FR:" & vbCrLf & "Stop CoreBridge harness" & vbCrLf & vbCrLf & _
        "EN:" & vbCrLf & "CoreBridge harness stop"
    DiagnosticProducersHarness_Trace "03 corebridge stop produced"

    CalcBridge_AddConsoleMessage messages, "INFO", _
        "FR:" & vbCrLf & "Info CoreBridge harness" & vbCrLf & vbCrLf & _
        "EN:" & vbCrLf & "CoreBridge harness info"
    DiagnosticProducersHarness_Trace "04 corebridge info produced"

    eventHash = BuildPlanningEventHash( _
        "WARNING", "CONSTRAINT_PARENT_IGNORED", _
        "Contrainte active ignoree sur une tache parent", _
        "Active constraint ignored on a summary task", _
        "diagnostic producer harness", _
        "diagnostic producer harness", _
        "DiagnosticProducersHarness", _
        "CONSTRAINTS", _
        "tbl_CONSTRAINTS", _
        "DP-CONSTRAINT-001", _
        "1.1", _
        "Harness constraint task")

    LogPlanningEvent _
        "WARNING", _
        "CONSTRAINT_PARENT_IGNORED", _
        eventHash, _
        "Contrainte active ignoree sur une tache parent", _
        "Active constraint ignored on a summary task", _
        "diagnostic producer harness", _
        "diagnostic producer harness", _
        "DiagnosticProducersHarness", _
        "CONSTRAINTS", _
        "tbl_CONSTRAINTS", _
        "DP-CONSTRAINT-001", _
        "1.1", _
        "Harness constraint task", _
        False

    CalcBridge_AddConsoleMessage messages, "WARNING", _
        "FR:" & vbCrLf & "Contrainte active ignoree sur une tache parent" & vbCrLf & vbCrLf & _
        "EN:" & vbCrLf & "Active constraint ignored on a summary task", _
        True, _
        "CONSTRAINT_PARENT_IGNORED", _
        eventHash
    DiagnosticProducersHarness_Trace "05 constraints warning produced"

    eventHash = BuildPlanningEventHash( _
        "WARNING", "SCURVE_MISSING_WEIGHT", _
        "Poids manquant sur certaines taches feuilles", _
        "Missing weight on some leaf tasks", _
        "diagnostic producer harness", _
        "diagnostic producer harness", _
        "DiagnosticProducersHarness", _
        "SCURVE", _
        "tbl_SCURVE", _
        "DP-SCURVE-001", _
        "2.1", _
        "Harness scurve task")
    ackToken = BuildPlanningWarningAckToken("SCURVE_MISSING_WEIGHT", eventHash)

    LogPlanningEvent _
        "WARNING", _
        "SCURVE_MISSING_WEIGHT", _
        eventHash, _
        "Poids manquant sur certaines taches feuilles", _
        "Missing weight on some leaf tasks", _
        "diagnostic producer harness", _
        "diagnostic producer harness", _
        "DiagnosticProducersHarness", _
        "SCURVE", _
        "tbl_SCURVE", _
        "DP-SCURVE-001", _
        "2.1", _
        "Harness scurve task", _
        False

    CalcBridge_AddConsoleMessage messages, "WARNING", _
        "FR:" & vbCrLf & "Poids manquant sur certaines taches feuilles" & vbCrLf & vbCrLf & _
        "EN:" & vbCrLf & "Missing weight on some leaf tasks", _
        True, _
        "SCURVE_MISSING_WEIGHT", _
        eventHash, _
        ackToken
    DiagnosticProducersHarness_Trace "06 scurve warning produced"

    Set mapCore = CreateObject("Scripting.Dictionary")
    mapCore("ID") = 1
    mapCore("WBS") = 2
    mapCore("Task Name") = 3
    mapCore("Error flag") = 4
    mapCore("ErrorMsg") = 5

    coreData(1, 1) = "CB-MISSING-001"
    coreData(1, 2) = "3.1"
    coreData(1, 3) = "Harness missing predecessor"
    coreData(1, 4) = "ERROR"
    coreData(1, 5) = "Missing predecessor"

    coreData(2, 1) = "CB-CONSTRAINT-001"
    coreData(2, 2) = "3.2"
    coreData(2, 3) = "Harness constraint stop"
    coreData(2, 4) = "ERROR"
    coreData(2, 5) = "Forecast Finish violates finish constraints"

    coreData(3, 1) = "CB-CONSTRAINT-FALLBACK-001"
    coreData(3, 2) = "3.4"
    coreData(3, 3) = "Harness constraint fallback"
    coreData(3, 4) = "ERROR"
    coreData(3, 5) = "Type de contrainte debut non reconnu"

    coreData(4, 1) = "CB-CYCLE-001"
    coreData(4, 2) = "3.5"
    coreData(4, 3) = "Harness cycle detail"
    coreData(4, 4) = "ERROR"
    coreData(4, 5) = "Cycle detected" & vbCrLf & "FR:" & vbCrLf & "Boucle harness detail" & vbCrLf & vbCrLf & "EN:" & vbCrLf & "Harness cycle detail"

    Set constraintDiagnostics = CreateObject("Scripting.Dictionary")
    Set diag = CreateObject("Scripting.Dictionary")
    diag("TaskID") = "CB-CONSTRAINT-001"
    diag("ConstraintType") = "Finish No Later Than"
    diag("CheckedField") = "Forecast Finish"
    diag("ExpectedOperator") = "<="
    diag("ConstraintDate") = DateSerial(2026, 1, 15)
    diag("CheckedValue") = DateSerial(2026, 1, 20)
    diag("AllowedValue") = DateSerial(2026, 1, 15)
    diag("CalculatedStart") = DateSerial(2026, 1, 10)
    diag("CalculatedFinish") = DateSerial(2026, 1, 20)
    Set constraintDiagnostics("CB-CONSTRAINT-001") = diag

    CalcBridge_AppendCoreErrorMessagesFromData messages, coreData, mapCore, Nothing, "PROD", Nothing, constraintDiagnostics
    DiagnosticProducersHarness_Trace "07 corebridge grouped and structured stop produced"

    testData(1, 1) = "CB-TEST-001"
    testData(1, 2) = "3.3"
    testData(1, 3) = "Harness test dependency stop"
    testData(1, 4) = "ERROR"
    testData(1, 5) = "Forecast Start violates dependencies"

    Set dependencyDiagnostics = CreateObject("Scripting.Dictionary")
    Set diag = CreateObject("Scripting.Dictionary")
    diag("TaskID") = "CB-TEST-001"
    diag("BlockingPredecessorID") = "CB-PRED-001"
    diag("BlockingLinkType") = "FS"
    diag("BlockingLag") = 2#
    diag("BlockingPredecessorDateKind") = "FINISH"
    diag("BlockingPredecessorDate") = DateSerial(2026, 2, 10)
    diag("MinimumAllowedStart") = DateSerial(2026, 2, 12)
    diag("RequestedStart") = DateSerial(2026, 2, 8)
    Set dependencyDiagnostics("CB-TEST-001") = diag

    CalcBridge_AppendCoreErrorMessagesFromData messages, testData, mapCore, Nothing, "TEST", dependencyDiagnostics
    DiagnosticProducersHarness_Trace "08 corebridge dependency stop produced"

    historyBefore = DiagnosticProducersHarness_TableRowCount("EVENT_HISTORY", "tbl_EVENT_HISTORY")
    PlanningConsolePolicy_EnableNonInteractive capturePath, "DiagnosticProducersHarness"
    CalcBridge_ShowPlanningConsole messages
    DiagnosticProducersHarness_Trace "09 console show returned"
    DiagnosticProducersHarness_Assert PlanningConsolePolicy_GetCapturedMessageCount() >= 3, "console captured diagnostics"
    PlanningConsolePolicy_DisableNonInteractive
    historyAfter = DiagnosticProducersHarness_TableRowCount("EVENT_HISTORY", "tbl_EVENT_HISTORY")
    DiagnosticProducersHarness_Assert historyAfter >= historyBefore, "event history remained available"

    DiagnosticProducersHarness_Assert DiagnosticProducersHarness_EventTypeExists("CONSTRAINT_PARENT_IGNORED"), "constraints event history present"
    DiagnosticProducersHarness_Assert DiagnosticProducersHarness_EventTypeExists("SCURVE_MISSING_WEIGHT"), "scurve event history present"
    DiagnosticProducersHarness_Assert DiagnosticProducersHarness_EventTypeExists("CONSOLE_STOP"), "corebridge stop history present"
    DiagnosticProducersHarness_Assert DiagnosticProducersHarness_EventTypeExists("CONSOLE_INFO"), "corebridge info history present"
    DiagnosticProducersHarness_Assert DiagnosticProducersHarness_CaptureHasText(capturePath, "Missing predecessor"), "corebridge grouped stop present"
    DiagnosticProducersHarness_Assert DiagnosticProducersHarness_CaptureHasText(capturePath, "Constraint cannot be met"), "corebridge structured constraint stop present"
    DiagnosticProducersHarness_Assert DiagnosticProducersHarness_CaptureHasText(capturePath, "Type de contrainte"), "corebridge constraint fallback stop present"
    DiagnosticProducersHarness_Assert DiagnosticProducersHarness_CaptureHasText(capturePath, "Harness cycle detail"), "corebridge cycle detail stop present"
    DiagnosticProducersHarness_Assert DiagnosticProducersHarness_CaptureHasText(capturePath, "Test Start impossible"), "corebridge dependency stop present"

    DiagnosticProducersHarness_WriteMatrix matrixPath, capturePath
    DiagnosticProducersHarness_Trace "10 matrix written"

    PlanningConsolePolicy_DisableNonInteractive
    DiagnosticProducersHarness_Trace "11 pass"
    DiagnosticProducersHarness_Smoke = "PASS"
    Exit Function

Fail:
    On Error Resume Next
    DiagnosticProducersHarness_Trace "FAIL " & Err.Number & " " & Err.Description
    PlanningConsolePolicy_DisableNonInteractive
    DiagnosticProducersHarness_Smoke = "FAIL: " & Err.Description

End Function

'------------------------------------------------------------------------------
' FR: Verifie le contrat Diagnostic Producers Harness Write Matrix et signale toute divergence au harnais.
' EN: Verifies the Diagnostic Producers Harness Write Matrix contract and reports any divergence to the harness.
'------------------------------------------------------------------------------

Private Sub DiagnosticProducersHarness_WriteMatrix( _
    ByVal matrixPath As String, _
    ByVal capturePath As String)

    Dim fileNo As Integer

    fileNo = FreeFile
    Open matrixPath For Output As #fileNo
    Print #fileNo, "Producer" & vbTab & "Message" & vbTab & "Severity" & vbTab & "Destination" & vbTab & "EventHistoryPresent" & vbTab & "ConsolePresent"
    DiagnosticProducersHarness_WriteMatrixRow fileNo, "CoreBridge", "CONSOLE_STOP", "STOP", "Console + EventHistory", DiagnosticProducersHarness_EventTypeExists("CONSOLE_STOP"), DiagnosticProducersHarness_CaptureHasSeverity(capturePath, "STOP")
    DiagnosticProducersHarness_WriteMatrixRow fileNo, "CoreBridge", "CONSOLE_INFO", "INFO", "Console + EventHistory", DiagnosticProducersHarness_EventTypeExists("CONSOLE_INFO"), DiagnosticProducersHarness_CaptureHasSeverity(capturePath, "INFO")
    DiagnosticProducersHarness_WriteMatrixRow fileNo, "CoreBridge", "CORE_GROUPED_STOP", "STOP", "Grouped core error formatting", DiagnosticProducersHarness_EventTypeExists("CONSOLE_STOP"), DiagnosticProducersHarness_CaptureHasText(capturePath, "Missing predecessor")
    DiagnosticProducersHarness_WriteMatrixRow fileNo, "CoreBridge", "CORE_CONSTRAINT_STRUCTURED_STOP", "STOP", "Structured constraint diagnostic formatting", DiagnosticProducersHarness_EventTypeExists("CONSOLE_STOP"), DiagnosticProducersHarness_CaptureHasText(capturePath, "Constraint cannot be met")
    DiagnosticProducersHarness_WriteMatrixRow fileNo, "CoreBridge", "CORE_CONSTRAINT_FALLBACK_STOP", "STOP", "PM-facing constraint fallback formatting", DiagnosticProducersHarness_EventTypeExists("CONSOLE_STOP"), DiagnosticProducersHarness_CaptureHasText(capturePath, "Type de contrainte")
    DiagnosticProducersHarness_WriteMatrixRow fileNo, "CoreBridge", "CORE_CYCLE_DETAIL_STOP", "STOP", "Cycle detail formatting", DiagnosticProducersHarness_EventTypeExists("CONSOLE_STOP"), DiagnosticProducersHarness_CaptureHasText(capturePath, "Harness cycle detail")
    DiagnosticProducersHarness_WriteMatrixRow fileNo, "CoreBridge", "CORE_DEPENDENCY_TEST_STOP", "STOP", "TEST dependency diagnostic formatting", DiagnosticProducersHarness_EventTypeExists("CONSOLE_STOP"), DiagnosticProducersHarness_CaptureHasText(capturePath, "Test Start impossible")
    DiagnosticProducersHarness_WriteMatrixRow fileNo, "Constraints", "CONSTRAINT_PARENT_IGNORED", "WARNING", "Console + EventHistory + ACK-ready", DiagnosticProducersHarness_EventTypeExists("CONSTRAINT_PARENT_IGNORED"), DiagnosticProducersHarness_CaptureHasSeverity(capturePath, "WARNING")
    DiagnosticProducersHarness_WriteMatrixRow fileNo, "S-Curve", "SCURVE_MISSING_WEIGHT", "WARNING", "Console + EventHistory + ACK token", DiagnosticProducersHarness_EventTypeExists("SCURVE_MISSING_WEIGHT"), DiagnosticProducersHarness_CaptureHasText(capturePath, "Missing weight")
    Close #fileNo

End Sub

'------------------------------------------------------------------------------
' FR: Verifie le contrat Diagnostic Producers Harness Write Matrix Row et signale toute divergence au harnais.
' EN: Verifies the Diagnostic Producers Harness Write Matrix Row contract and reports any divergence to the harness.
'------------------------------------------------------------------------------

Private Sub DiagnosticProducersHarness_WriteMatrixRow( _
    ByVal fileNo As Integer, _
    ByVal producerName As String, _
    ByVal messageName As String, _
    ByVal severityName As String, _
    ByVal destinationName As String, _
    ByVal eventHistoryPresent As Boolean, _
    ByVal consolePresent As Boolean)

    Print #fileNo, _
        DiagnosticProducersHarness_Tsv(producerName) & vbTab & _
        DiagnosticProducersHarness_Tsv(messageName) & vbTab & _
        DiagnosticProducersHarness_Tsv(severityName) & vbTab & _
        DiagnosticProducersHarness_Tsv(destinationName) & vbTab & _
        CStr(eventHistoryPresent) & vbTab & _
        CStr(consolePresent)

End Sub

'------------------------------------------------------------------------------
' FR: Verifie le contrat Diagnostic Producers Harness Event Type Exists et signale toute divergence au harnais.
' EN: Verifies the Diagnostic Producers Harness Event Type Exists contract and reports any divergence to the harness.
'------------------------------------------------------------------------------

Private Function DiagnosticProducersHarness_EventTypeExists(ByVal eventType As String) As Boolean

    Dim tbl As ListObject
    Dim colIdx As Long
    Dim r As Long

    On Error GoTo SafeExit
    Set tbl = ThisWorkbook.Worksheets("CALC_ALARM").ListObjects("tbl_CALC_ALARM")
    If tbl.DataBodyRange Is Nothing Then Exit Function
    colIdx = tbl.ListColumns("Event Type").Index

    For r = 1 To tbl.ListRows.Count
        If UCase$(Trim$(CStr(tbl.DataBodyRange.Cells(r, colIdx).value))) = UCase$(Trim$(eventType)) Then
            DiagnosticProducersHarness_EventTypeExists = True
            Exit Function
        End If
    Next r

SafeExit:
End Function

'------------------------------------------------------------------------------
' FR: Verifie le contrat Diagnostic Producers Harness Capture Has Severity et signale toute divergence au harnais.
' EN: Verifies the Diagnostic Producers Harness Capture Has Severity contract and reports any divergence to the harness.
'------------------------------------------------------------------------------

Private Function DiagnosticProducersHarness_CaptureHasSeverity( _
    ByVal capturePath As String, _
    ByVal severityName As String) As Boolean

    DiagnosticProducersHarness_CaptureHasSeverity = _
        DiagnosticProducersHarness_CaptureHasText(capturePath, vbTab & UCase$(Trim$(severityName)) & vbTab)

End Function

'------------------------------------------------------------------------------
' FR: Verifie le contrat Diagnostic Producers Harness Capture Has Text et signale toute divergence au harnais.
' EN: Verifies the Diagnostic Producers Harness Capture Has Text contract and reports any divergence to the harness.
'------------------------------------------------------------------------------

Private Function DiagnosticProducersHarness_CaptureHasText( _
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
            DiagnosticProducersHarness_CaptureHasText = True
            Exit Do
        End If
    Loop
    Close #fileNo

SafeExit:
    On Error Resume Next
    Close #fileNo
End Function

'------------------------------------------------------------------------------
' FR: Verifie le contrat Diagnostic Producers Harness Table Row Count et signale toute divergence au harnais.
' EN: Verifies the Diagnostic Producers Harness Table Row Count contract and reports any divergence to the harness.
'------------------------------------------------------------------------------

Private Function DiagnosticProducersHarness_TableRowCount( _
    ByVal sheetName As String, _
    ByVal tableName As String) As Long

    Dim tbl As ListObject

    On Error GoTo SafeExit
    Set tbl = ThisWorkbook.Worksheets(sheetName).ListObjects(tableName)
    If tbl.DataBodyRange Is Nothing Then Exit Function
    DiagnosticProducersHarness_TableRowCount = tbl.ListRows.Count

SafeExit:
End Function

'------------------------------------------------------------------------------
' FR: Verifie le contrat Diagnostic Producers Harness Assert et signale toute divergence au harnais.
' EN: Verifies the Diagnostic Producers Harness Assert contract and reports any divergence to the harness.
'------------------------------------------------------------------------------

Private Sub DiagnosticProducersHarness_Assert(ByVal condition As Boolean, ByVal messageText As String)

    If Not condition Then
        DiagnosticProducersHarness_Trace "ASSERT FAIL " & messageText
        Err.Raise vbObjectError + 9371, "DiagnosticProducersHarness", messageText
    End If

End Sub

'------------------------------------------------------------------------------
' FR: Verifie le contrat Diagnostic Producers Harness Trace et signale toute divergence au harnais.
' EN: Verifies the Diagnostic Producers Harness Trace contract and reports any divergence to the harness.
'------------------------------------------------------------------------------

Private Sub DiagnosticProducersHarness_Trace(ByVal messageText As String)

    Dim fileNo As Integer

    On Error Resume Next
    If Trim$(gDiagnosticProducersTracePath) = "" Then Exit Sub
    fileNo = FreeFile
    Open gDiagnosticProducersTracePath For Append As #fileNo
    Print #fileNo, Format$(Now, "yyyy-mm-dd hh:nn:ss") & " | " & messageText
    Close #fileNo

End Sub

'------------------------------------------------------------------------------
' FR: Verifie le contrat Diagnostic Producers Harness Delete File et signale toute divergence au harnais.
' EN: Verifies the Diagnostic Producers Harness Delete File contract and reports any divergence to the harness.
'------------------------------------------------------------------------------

Private Sub DiagnosticProducersHarness_DeleteFile(ByVal filePath As String)

    On Error Resume Next
    If Len(Dir$(filePath)) > 0 Then Kill filePath

End Sub

'------------------------------------------------------------------------------
' FR: Verifie le contrat Diagnostic Producers Harness Tsv et signale toute divergence au harnais.
' EN: Verifies the Diagnostic Producers Harness Tsv contract and reports any divergence to the harness.
'------------------------------------------------------------------------------

Private Function DiagnosticProducersHarness_Tsv(ByVal value As String) As String

    DiagnosticProducersHarness_Tsv = Replace(Replace(Replace(value, vbTab, " "), vbCr, "\r"), vbLf, "\n")

End Function
