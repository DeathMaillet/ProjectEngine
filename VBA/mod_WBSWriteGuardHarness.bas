Attribute VB_Name = "mod_WBSWriteGuardHarness"
Option Explicit

'===============================================================================
' MODULE : mod_WBSWriteGuardHarness
' DOMAINE / DOMAIN : Validation Harnesses
'
' FR
' Harnais de preuve du contrat WBS Write Guard sur des copies de test.
' N'appartient a aucun workflow produit et ne doit pas etre appele en usage normal.
'
' EN
' Proof harness for the WBS Write Guard contract on test copies.
' Is not production workflow code and must not run during normal use.
'
' CONTRATS / CONTRACTS : WBSWriteGuardHarness_Smoke
' CALLBACKS EXTERNES / EXTERNAL CALLBACKS : Aucun / None
'===============================================================================


Private gWBSGuardHarnessTracePath As String

'------------------------------------------------------------------------------
' FR: Caracterise le cycle de vie du guard WBS sur une copie temporaire.
' EN: Characterizes the WBS write guard lifecycle on a temporary workbook copy.
'------------------------------------------------------------------------------
Public Function WBSWriteGuardHarness_Smoke(ByVal matrixPath As String) As String

    Dim wsWBS As Worksheet
    Dim tblWBS As ListObject
    Dim messages As Collection
    Dim fullWriterLeakSource As String
    Dim partialErrorObservedSource As String
    Dim analyticsErrorObservedSource As String
    Dim criticalPathErrorLeakSource As String
    Dim runtimeLeakSource As String
    Dim scenarioForkLeakSource As String

    On Error GoTo Fail

    gWBSGuardHarnessTracePath = matrixPath & ".trace.txt"
    WBSWriteGuardHarness_DeleteFile gWBSGuardHarnessTracePath
    WBSWriteGuardHarness_DeleteFile matrixPath
    WBSWriteGuardHarness_DeleteFile matrixPath & ".messages.tsv"
    WBSWriteGuardHarness_Trace "01 enter"

    PlanningConsolePolicy_EnableNonInteractive matrixPath & ".messages.tsv", "WBSWriteGuardHarness"
    WBSWriteGuardHarness_Reset
    AppEvents_EnsureInitialized

    Set wsWBS = ThisWorkbook.Worksheets("WBS")
    Set tblWBS = wsWBS.ListObjects("tbl_WBS")
    Set messages = New Collection

    WBSWriteGuardHarness_Assert Not tblWBS.DataBodyRange Is Nothing, "tbl_WBS has rows"

    WBSWriteGuardHarness_TestPrimitiveLifecycle
    WBSWriteGuardHarness_Trace "02 primitive lifecycle ok"

    WBSWriteGuardHarness_TestNestedLifecycle
    WBSWriteGuardHarness_Trace "03 nested lifecycle ok"

    WBSWriteGuardHarness_TestTokenOwnership
    WBSWriteGuardHarness_Trace "03b token ownership ok"

    WBSWriteGuardHarness_TestErrorCleanup
    WBSWriteGuardHarness_Trace "04 controlled error cleanup ok"

    WBSWriteGuardHarness_TestWorksheetEvent wsWBS, tblWBS
    WBSWriteGuardHarness_Trace "05 worksheet event contract ok"

    WBSWriteGuardHarness_TestBalancedWriters tblWBS, messages
    WBSWriteGuardHarness_Trace "06 balanced writers ok"

    criticalPathErrorLeakSource = WBSWriteGuardHarness_TestWriterErrorPaths(wsWBS, tblWBS, messages)
    WBSWriteGuardHarness_Assert criticalPathErrorLeakSource = "INACTIVE", _
        "critical path error scope closed"
    WBSWriteGuardHarness_Trace "07 writer error paths pass"
    WBSWriteGuardHarness_Reset

    fullWriterLeakSource = WBSWriteGuardHarness_CharacterizeFullWriterLeak
    WBSWriteGuardHarness_Assert fullWriterLeakSource = "INACTIVE", _
        "full writer final inactive"
    WBSWriteGuardHarness_Trace "08 full writer lifecycle pass"
    WBSWriteGuardHarness_Reset

    partialErrorObservedSource = WBSWriteGuardHarness_CharacterizePartialPreBeginError(tblWBS)
    WBSWriteGuardHarness_Assert partialErrorObservedSource = "PARENT_RESTORED", _
        "partial pre-begin error preserves parent scope"
    WBSWriteGuardHarness_Trace "09 partial error parent restoration pass"
    WBSWriteGuardHarness_Reset

    analyticsErrorObservedSource = WBSWriteGuardHarness_CharacterizeAnalyticsPreBeginError(tblWBS)
    WBSWriteGuardHarness_Assert analyticsErrorObservedSource = "PARENT_RESTORED", _
        "analytics pre-begin error preserves parent scope"
    WBSWriteGuardHarness_Trace "10 analytics error parent restoration pass"
    WBSWriteGuardHarness_Reset

    runtimeLeakSource = WBSWriteGuardHarness_CharacterizeRuntimeLeak
    WBSWriteGuardHarness_Assert runtimeLeakSource = "INACTIVE", _
        "forced runtime final inactive"
    WBSWriteGuardHarness_Trace "11 runtime lifecycle pass"
    WBSWriteGuardHarness_Reset

    scenarioForkLeakSource = WBSWriteGuardHarness_CharacterizeScenarioForkNesting
    WBSWriteGuardHarness_Assert scenarioForkLeakSource = "INACTIVE", _
        "scenario fork forced-update final inactive"
    WBSWriteGuardHarness_Trace "12 scenario fork nesting pass"
    WBSWriteGuardHarness_Reset

    WBSWriteGuardHarness_WriteMatrix matrixPath, fullWriterLeakSource, _
        partialErrorObservedSource, analyticsErrorObservedSource, _
        criticalPathErrorLeakSource, runtimeLeakSource, scenarioForkLeakSource
    WBSWriteGuardHarness_Trace "13 matrix written"

    PlanningConsolePolicy_DisableNonInteractive
    WBSWriteGuardHarness_Smoke = "PASS"
    Exit Function

Fail:
    On Error Resume Next
    WBSWriteGuardHarness_Trace "FAIL " & Err.Number & " " & Err.Description
    If IsMacroRunActive() Then EndMacroRun
    Application.EnableEvents = True
    WBSWriteGuardHarness_Reset
    PlanningConsolePolicy_DisableNonInteractive
    WBSWriteGuardHarness_Smoke = "FAIL: " & Err.Description

End Function

'------------------------------------------------------------------------------
' FR: Verifie le contrat WBS Write Guard Harness Test Primitive Lifecycle et signale toute divergence au harnais.
' EN: Verifies the WBS Write Guard Harness Test Primitive Lifecycle contract and reports any divergence to the harness.
'------------------------------------------------------------------------------

Private Sub WBSWriteGuardHarness_TestPrimitiveLifecycle()

    WBSWriteGuardHarness_Reset
    BeginAuthorizedWBSWrite "HarnessPrimitive", Array("Calculated Start")

    WBSWriteGuardHarness_Assert IsAuthorizedWBSWriteActive(), "primitive scope active"
    WBSWriteGuardHarness_Assert GetAuthorizedWBSWriteSource() = "HarnessPrimitive", "primitive source"
    WBSWriteGuardHarness_Assert IsAuthorizedWBSColumn("Calculated Start"), "primitive allowed column"
    WBSWriteGuardHarness_Assert Not IsAuthorizedWBSColumn("Deadline Float"), "primitive rejected column"

    EndAuthorizedWBSWrite
    WBSWriteGuardHarness_Assert Not IsAuthorizedWBSWriteActive(), "primitive final inactive"

End Sub

'------------------------------------------------------------------------------
' FR: Verifie le contrat WBS Write Guard Harness Test Writer Error Paths et signale toute divergence au harnais.
' EN: Verifies the WBS Write Guard Harness Test Writer Error Paths contract and reports any divergence to the harness.
'------------------------------------------------------------------------------

Private Function WBSWriteGuardHarness_TestWriterErrorPaths( _
    ByVal wsWBS As Worksheet, _
    ByVal tblWBS As ListObject, _
    ByVal messages As Collection) As String

    Dim impactedIds As Object
    Dim firstId As String
    Dim criticalPathErrorCaught As Boolean
    Dim originalErrorNumber As Long
    Dim originalErrorDescription As String

    On Error GoTo Fail

    firstId = Trim$(CStr(tblWBS.ListColumns("ID").DataBodyRange.Cells(1, 1).Value))
    Set impactedIds = CreateObject("Scripting.Dictionary")
    impactedIds(firstId) = True

    wsWBS.Protect

    WBSWriteGuardHarness_Reset
    Push_Calculated_Back_To_WBS_Partial impactedIds
    WBSWriteGuardHarness_Assert Not IsAuthorizedWBSWriteActive(), _
        "partial writer after-begin error cleanup"

    RestoreWBSFormulaColumns tblWBS
    WBSWriteGuardHarness_Assert Not IsAuthorizedWBSWriteActive(), _
        "formula writer after-begin error cleanup"

    Compute_And_Push_Variances messages
    WBSWriteGuardHarness_Assert Not IsAuthorizedWBSWriteActive(), _
        "variance writer after-begin error cleanup"

    WBSWriteGuardHarness_Reset
    criticalPathErrorCaught = WBSWriteGuardHarness_InvokeCriticalPathExpectError()
    WBSWriteGuardHarness_Assert criticalPathErrorCaught, "critical path protected-sheet error caught"
    WBSWriteGuardHarness_Assert Not IsAuthorizedWBSWriteActive(), _
        "critical path error closes owned scope"
    WBSWriteGuardHarness_TestWriterErrorPaths = "INACTIVE"

    wsWBS.Unprotect
    Exit Function

Fail:
    originalErrorNumber = Err.Number
    originalErrorDescription = Err.Description
    On Error Resume Next
    wsWBS.Unprotect
    WBSWriteGuardHarness_Reset
    On Error GoTo 0
    Err.Raise originalErrorNumber, "WBSWriteGuardHarness_TestWriterErrorPaths", originalErrorDescription

End Function

'------------------------------------------------------------------------------
' FR: Verifie le contrat WBS Write Guard Harness Invoke Critical Path Expect Error et signale toute divergence au harnais.
' EN: Verifies the WBS Write Guard Harness Invoke Critical Path Expect Error contract and reports any divergence to the harness.
'------------------------------------------------------------------------------

Private Function WBSWriteGuardHarness_InvokeCriticalPathExpectError() As Boolean

    On Error GoTo ExpectedError
    Push_CriticalPathREX_Back_To_WBS
    Exit Function

ExpectedError:
    WBSWriteGuardHarness_InvokeCriticalPathExpectError = (Err.Number <> 0)
    Err.Clear

End Function

'------------------------------------------------------------------------------
' FR: Verifie le contrat WBS Write Guard Harness Test Nested Lifecycle et signale toute divergence au harnais.
' EN: Verifies the WBS Write Guard Harness Test Nested Lifecycle contract and reports any divergence to the harness.
'------------------------------------------------------------------------------

Private Sub WBSWriteGuardHarness_TestNestedLifecycle()

    WBSWriteGuardHarness_Reset
    BeginAuthorizedWBSWrite "HarnessOuter", Array("Calculated Start")
    BeginAuthorizedWBSWrite "HarnessInner", Array("Deadline Float")

    WBSWriteGuardHarness_Assert GetAuthorizedWBSWriteSource() = "HarnessInner", "inner source active"
    WBSWriteGuardHarness_Assert IsAuthorizedWBSColumn("Deadline Float"), "inner column active"
    WBSWriteGuardHarness_Assert Not IsAuthorizedWBSColumn("Calculated Start"), "outer column hidden by inner"

    EndAuthorizedWBSWrite

    WBSWriteGuardHarness_Assert IsAuthorizedWBSWriteActive(), "outer restored active"
    WBSWriteGuardHarness_Assert GetAuthorizedWBSWriteSource() = "HarnessOuter", "outer source restored"
    WBSWriteGuardHarness_Assert IsAuthorizedWBSColumn("Calculated Start"), "outer column restored"
    WBSWriteGuardHarness_Assert Not IsAuthorizedWBSColumn("Deadline Float"), "inner column removed"

    EndAuthorizedWBSWrite
    WBSWriteGuardHarness_Assert Not IsAuthorizedWBSWriteActive(), "nested final inactive"

End Sub

'------------------------------------------------------------------------------
' FR: Verifie le contrat WBS Write Guard Harness Test Token Ownership et signale toute divergence au harnais.
' EN: Verifies the WBS Write Guard Harness Test Token Ownership contract and reports any divergence to the harness.
'------------------------------------------------------------------------------

Private Sub WBSWriteGuardHarness_TestTokenOwnership()

    Dim outerToken As Long
    Dim innerToken As Long
    Dim mismatchRejected As Boolean

    WBSWriteGuardHarness_Reset
    outerToken = OpenAuthorizedWBSWriteScope("HarnessTokenOuter", Array("Calculated Start"))
    innerToken = OpenAuthorizedWBSWriteScope("HarnessTokenInner", Array("Deadline Float"))

    On Error Resume Next
    CloseAuthorizedWBSWriteScope outerToken
    mismatchRejected = (Err.Number <> 0)
    Err.Clear
    On Error GoTo 0

    WBSWriteGuardHarness_Assert mismatchRejected, "out-of-order token close rejected"
    WBSWriteGuardHarness_Assert GetAuthorizedWBSWriteSource() = "HarnessTokenInner", _
        "rejected close preserves inner scope"

    CloseAuthorizedWBSWriteScope innerToken
    WBSWriteGuardHarness_Assert GetAuthorizedWBSWriteSource() = "HarnessTokenOuter", _
        "owned inner close restores outer scope"

    CloseAuthorizedWBSWriteScope outerToken
    WBSWriteGuardHarness_Assert Not IsAuthorizedWBSWriteActive(), _
        "token ownership final inactive"

End Sub

'------------------------------------------------------------------------------
' FR: Verifie le contrat WBS Write Guard Harness Test Error Cleanup et signale toute divergence au harnais.
' EN: Verifies the WBS Write Guard Harness Test Error Cleanup contract and reports any divergence to the harness.
'------------------------------------------------------------------------------

Private Sub WBSWriteGuardHarness_TestErrorCleanup()

    WBSWriteGuardHarness_Reset
    WBSWriteGuardHarness_Assert WBSWriteGuardHarness_RaiseInsideBalancedScope(), _
        "controlled error was caught"
    WBSWriteGuardHarness_Assert Not IsAuthorizedWBSWriteActive(), _
        "controlled error cleanup final inactive"

End Sub

'------------------------------------------------------------------------------
' FR: Verifie le contrat WBS Write Guard Harness Raise Inside Balanced Scope et signale toute divergence au harnais.
' EN: Verifies the WBS Write Guard Harness Raise Inside Balanced Scope contract and reports any divergence to the harness.
'------------------------------------------------------------------------------

Private Function WBSWriteGuardHarness_RaiseInsideBalancedScope() As Boolean

    Dim scopeStarted As Boolean
    Dim errorWasRaised As Boolean

    On Error GoTo ExpectedError

    BeginAuthorizedWBSWrite "HarnessError", Array("Calculated Start")
    scopeStarted = True
    Err.Raise vbObjectError + 9610, "WBSWriteGuardHarness", "Expected harness error"

ExpectedError:
    errorWasRaised = (Err.Number <> 0)
    On Error Resume Next
    If scopeStarted Then EndAuthorizedWBSWrite
    scopeStarted = False
    WBSWriteGuardHarness_RaiseInsideBalancedScope = errorWasRaised
    Err.Clear
    On Error GoTo 0

End Function

'------------------------------------------------------------------------------
' FR: Verifie le contrat WBS Write Guard Harness Test Worksheet Event et signale toute divergence au harnais.
' EN: Verifies the WBS Write Guard Harness Test Worksheet Event contract and reports any divergence to the harness.
'------------------------------------------------------------------------------

Private Sub WBSWriteGuardHarness_TestWorksheetEvent( _
    ByVal wsWBS As Worksheet, _
    ByVal tblWBS As ListObject)

    Dim targetCell As Range

    Set targetCell = tblWBS.ListColumns("Calculated Start").DataBodyRange.Cells(1, 1)

    WBSWriteGuardHarness_Reset
    BeginMacroRun "WBSWriteGuardHarnessAuthorizedEvent"
    BeginAuthorizedWBSWrite "HarnessEvent", Array("Calculated Start")
    Handle_WBS_Change wsWBS, targetCell
    WBSWriteGuardHarness_Assert Not IsMacroAbortRequested(), "authorized worksheet event accepted"
    EndAuthorizedWBSWrite
    EndMacroRun
    Application.EnableEvents = True

    BeginMacroRun "WBSWriteGuardHarnessUnauthorizedEvent"
    Handle_WBS_Change wsWBS, targetCell
    WBSWriteGuardHarness_Assert IsMacroAbortRequested(), "unauthorized worksheet event rejected"
    EndMacroRun
    Application.EnableEvents = True
    WBSWriteGuardHarness_Assert Not IsAuthorizedWBSWriteActive(), "worksheet event final inactive"

End Sub

'------------------------------------------------------------------------------
' FR: Verifie le contrat WBS Write Guard Harness Test Balanced Writers et signale toute divergence au harnais.
' EN: Verifies the WBS Write Guard Harness Test Balanced Writers contract and reports any divergence to the harness.
'------------------------------------------------------------------------------

Private Sub WBSWriteGuardHarness_TestBalancedWriters( _
    ByVal tblWBS As ListObject, _
    ByVal messages As Collection)

    Dim impactedIds As Object
    Dim firstId As String

    firstId = Trim$(CStr(tblWBS.ListColumns("ID").DataBodyRange.Cells(1, 1).Value))
    WBSWriteGuardHarness_Assert firstId <> vbNullString, "first WBS row has ID"

    Set impactedIds = CreateObject("Scripting.Dictionary")
    impactedIds(firstId) = True

    WBSWriteGuardHarness_Reset
    Push_Calculated_Back_To_WBS_Partial impactedIds
    WBSWriteGuardHarness_Assert Not IsAuthorizedWBSWriteActive(), "partial writer final inactive"

    RestoreWBSFormulaColumns tblWBS
    WBSWriteGuardHarness_Assert Not IsAuthorizedWBSWriteActive(), "formula writer final inactive"

    Clear_Analytics_Outputs messages
    WBSWriteGuardHarness_Assert Not IsAuthorizedWBSWriteActive(), "analytics clear final inactive"

    Push_Analytics_Back_To_WBS
    WBSWriteGuardHarness_Assert Not IsAuthorizedWBSWriteActive(), "analytics push final inactive"

    Push_CriticalPathREX_Back_To_WBS
    WBSWriteGuardHarness_Assert Not IsAuthorizedWBSWriteActive(), "critical path REX push final inactive"

    Compute_And_Push_Variances messages
    WBSWriteGuardHarness_Assert Not IsAuthorizedWBSWriteActive(), "variance writer final inactive"

End Sub

'------------------------------------------------------------------------------
' FR: Verifie le contrat WBS Write Guard Harness Characterize Full Writer Leak et signale toute divergence au harnais.
' EN: Verifies the WBS Write Guard Harness Characterize Full Writer Leak contract and reports any divergence to the harness.
'------------------------------------------------------------------------------

Private Function WBSWriteGuardHarness_CharacterizeFullWriterLeak() As String

    WBSWriteGuardHarness_Reset
    Push_Calculated_Back_To_WBS

    If IsAuthorizedWBSWriteActive() Then
        WBSWriteGuardHarness_CharacterizeFullWriterLeak = GetAuthorizedWBSWriteSource()
    Else
        WBSWriteGuardHarness_CharacterizeFullWriterLeak = "INACTIVE"
    End If

End Function

'------------------------------------------------------------------------------
' FR: Verifie le contrat WBS Write Guard Harness Characterize Partial Pre Begin Error et signale toute divergence au harnais.
' EN: Verifies the WBS Write Guard Harness Characterize Partial Pre Begin Error contract and reports any divergence to the harness.
'------------------------------------------------------------------------------

Private Function WBSWriteGuardHarness_CharacterizePartialPreBeginError( _
    ByVal tblWBS As ListObject) As String

    Dim impactedIds As Object
    Dim targetColumn As ListColumn
    Dim originalName As String
    Dim firstId As String

    On Error GoTo Fail

    firstId = Trim$(CStr(tblWBS.ListColumns("ID").DataBodyRange.Cells(1, 1).Value))
    Set impactedIds = CreateObject("Scripting.Dictionary")
    impactedIds(firstId) = True

    Set targetColumn = tblWBS.ListColumns("Deadline Float")
    originalName = targetColumn.Name

    WBSWriteGuardHarness_Reset
    BeginAuthorizedWBSWrite "HarnessParent", Array("Calculated Start")

    Application.EnableEvents = False
    targetColumn.Name = "Deadline Float Harness Missing"
    Push_Calculated_Back_To_WBS_Partial impactedIds
    targetColumn.Name = originalName
    Application.EnableEvents = True

    WBSWriteGuardHarness_Assert IsAuthorizedWBSWriteActive(), _
        "partial pre-begin error keeps parent active"
    WBSWriteGuardHarness_Assert GetAuthorizedWBSWriteSource() = "HarnessParent", _
        "partial pre-begin error restores parent source"
    EndAuthorizedWBSWrite
    WBSWriteGuardHarness_CharacterizePartialPreBeginError = "PARENT_RESTORED"
    Exit Function

Fail:
    On Error Resume Next
    If Not targetColumn Is Nothing Then targetColumn.Name = originalName
    Application.EnableEvents = True
    WBSWriteGuardHarness_Reset
    Err.Raise Err.Number, "WBSWriteGuardHarness_CharacterizePartialPreBeginError", Err.Description

End Function

'------------------------------------------------------------------------------
' FR: Verifie le contrat WBS Write Guard Harness Characterize Runtime Leak et signale toute divergence au harnais.
' EN: Verifies the WBS Write Guard Harness Characterize Runtime Leak contract and reports any divergence to the harness.
'------------------------------------------------------------------------------

Private Function WBSWriteGuardHarness_CharacterizeRuntimeLeak() As String

    WBSWriteGuardHarness_Reset
    Run_Forced_Planning_Update

    If IsAuthorizedWBSWriteActive() Then
        WBSWriteGuardHarness_CharacterizeRuntimeLeak = GetAuthorizedWBSWriteSource()
    Else
        WBSWriteGuardHarness_CharacterizeRuntimeLeak = "INACTIVE"
    End If

End Function

'------------------------------------------------------------------------------
' FR: Verifie le contrat WBS Write Guard Harness Characterize Scenario Fork Nesting et signale toute divergence au harnais.
' EN: Verifies the WBS Write Guard Harness Characterize Scenario Fork Nesting contract and reports any divergence to the harness.
'------------------------------------------------------------------------------

Private Function WBSWriteGuardHarness_CharacterizeScenarioForkNesting() As String

    Dim parentToken As Long

    WBSWriteGuardHarness_Reset
    parentToken = OpenAuthorizedWBSWriteScope( _
        "ScenarioForkForcedUpdate", Array("Calculated Start", "Calculated Finish"))

    Run_Forced_Planning_Update
    WBSWriteGuardHarness_Assert IsAuthorizedWBSWriteActive(), _
        "scenario fork parent scope remains active"
    WBSWriteGuardHarness_Assert GetAuthorizedWBSWriteSource() = "ScenarioForkForcedUpdate", _
        "scenario fork parent source restored"

    CloseAuthorizedWBSWriteScope parentToken

    If IsAuthorizedWBSWriteActive() Then
        WBSWriteGuardHarness_CharacterizeScenarioForkNesting = GetAuthorizedWBSWriteSource()
    Else
        WBSWriteGuardHarness_CharacterizeScenarioForkNesting = "INACTIVE"
    End If

End Function

'------------------------------------------------------------------------------
' FR: Verifie le contrat WBS Write Guard Harness Characterize Analytics Pre Begin Error et signale toute divergence au harnais.
' EN: Verifies the WBS Write Guard Harness Characterize Analytics Pre Begin Error contract and reports any divergence to the harness.
'------------------------------------------------------------------------------

Private Function WBSWriteGuardHarness_CharacterizeAnalyticsPreBeginError( _
    ByVal tblWBS As ListObject) As String

    Dim targetColumn As ListColumn
    Dim originalName As String
    Dim originalErrorNumber As Long
    Dim originalErrorDescription As String

    On Error GoTo Fail

    Set targetColumn = tblWBS.ListColumns("ID")
    originalName = targetColumn.Name

    WBSWriteGuardHarness_Reset
    BeginAuthorizedWBSWrite "HarnessAnalyticsParent", Array("Calculated Start")

    Application.EnableEvents = False
    targetColumn.Name = "ID Harness Missing"
    On Error Resume Next
    Push_Analytics_Back_To_WBS
    Err.Clear
    On Error GoTo Fail
    targetColumn.Name = originalName
    Application.EnableEvents = True

    WBSWriteGuardHarness_Assert IsAuthorizedWBSWriteActive(), _
        "analytics pre-begin error keeps parent active"
    WBSWriteGuardHarness_Assert GetAuthorizedWBSWriteSource() = "HarnessAnalyticsParent", _
        "analytics pre-begin error restores parent source"
    EndAuthorizedWBSWrite
    WBSWriteGuardHarness_CharacterizeAnalyticsPreBeginError = "PARENT_RESTORED"
    Exit Function

Fail:
    originalErrorNumber = Err.Number
    originalErrorDescription = Err.Description
    On Error Resume Next
    If Not targetColumn Is Nothing Then targetColumn.Name = originalName
    Application.EnableEvents = True
    WBSWriteGuardHarness_Reset
    On Error GoTo 0
    Err.Raise originalErrorNumber, "WBSWriteGuardHarness_CharacterizeAnalyticsPreBeginError", originalErrorDescription

End Function

'------------------------------------------------------------------------------
' FR: Verifie le contrat WBS Write Guard Harness Write Matrix et signale toute divergence au harnais.
' EN: Verifies the WBS Write Guard Harness Write Matrix contract and reports any divergence to the harness.
'------------------------------------------------------------------------------

Private Sub WBSWriteGuardHarness_WriteMatrix( _
    ByVal matrixPath As String, _
    ByVal fullWriterLeakSource As String, _
    ByVal partialErrorObservedSource As String, _
    ByVal analyticsErrorObservedSource As String, _
    ByVal criticalPathErrorLeakSource As String, _
    ByVal runtimeLeakSource As String, _
    ByVal scenarioForkLeakSource As String)

    Dim fileNo As Integer

    fileNo = FreeFile
    Open matrixPath For Output As #fileNo
    Print #fileNo, "Check" & vbTab & "Path" & vbTab & "Expected" & vbTab & "Observed" & vbTab & "Classification"
    Print #fileNo, "Primitive begin/end" & vbTab & "Guard" & vbTab & "Inactive" & vbTab & "Inactive" & vbTab & "PASS"
    Print #fileNo, "Nested source restore" & vbTab & "Guard" & vbTab & "Outer restored" & vbTab & "Outer restored" & vbTab & "PASS"
    Print #fileNo, "Nested column restore" & vbTab & "Guard" & vbTab & "Outer permissions" & vbTab & "Outer permissions" & vbTab & "PASS"
    Print #fileNo, "Token ownership and LIFO rejection" & vbTab & "Guard" & vbTab & "Rejected without pop" & vbTab & "Rejected without pop" & vbTab & "PASS"
    Print #fileNo, "Controlled error cleanup" & vbTab & "Guard" & vbTab & "Inactive" & vbTab & "Inactive" & vbTab & "PASS"
    Print #fileNo, "Authorized Worksheet event" & vbTab & "Worksheet Events" & vbTab & "No abort" & vbTab & "No abort" & vbTab & "PASS"
    Print #fileNo, "Unauthorized Worksheet event" & vbTab & "Worksheet Events" & vbTab & "Abort" & vbTab & "Abort" & vbTab & "PASS"
    Print #fileNo, "Partial writer success" & vbTab & "Partial Writer" & vbTab & "Inactive" & vbTab & "Inactive" & vbTab & "PASS"
    Print #fileNo, "Formula writer success" & vbTab & "Formula Writer" & vbTab & "Inactive" & vbTab & "Inactive" & vbTab & "PASS"
    Print #fileNo, "Analytics clear success" & vbTab & "Analytics" & vbTab & "Inactive" & vbTab & "Inactive" & vbTab & "PASS"
    Print #fileNo, "Analytics push success" & vbTab & "Analytics" & vbTab & "Inactive" & vbTab & "Inactive" & vbTab & "PASS"
    Print #fileNo, "Critical Path REX success" & vbTab & "Analytics" & vbTab & "Inactive" & vbTab & "Inactive" & vbTab & "PASS"
    Print #fileNo, "Variance writer success" & vbTab & "Variances" & vbTab & "Inactive" & vbTab & "Inactive" & vbTab & "PASS"
    Print #fileNo, "Partial writer after-Begin error" & vbTab & "Partial Writer Error" & vbTab & "Inactive" & vbTab & "Inactive" & vbTab & "PASS"
    Print #fileNo, "Formula writer after-Begin error" & vbTab & "Formula Writer Error" & vbTab & "Inactive" & vbTab & "Inactive" & vbTab & "PASS"
    Print #fileNo, "Variance writer after-Begin error" & vbTab & "Variance Error" & vbTab & "Inactive" & vbTab & "Inactive" & vbTab & "PASS"
    Print #fileNo, "Critical Path REX after-Begin error" & vbTab & "Analytics Error" & vbTab & "Inactive" & vbTab & criticalPathErrorLeakSource & vbTab & "PASS"
    Print #fileNo, "Full writer final state" & vbTab & "Full Writer" & vbTab & "Inactive" & vbTab & fullWriterLeakSource & vbTab & "PASS"
    Print #fileNo, "Partial pre-begin error parent state" & vbTab & "Partial Writer Error" & vbTab & "Parent restored" & vbTab & partialErrorObservedSource & vbTab & "PASS"
    Print #fileNo, "Analytics pre-begin error parent state" & vbTab & "Analytics Error" & vbTab & "Parent restored" & vbTab & analyticsErrorObservedSource & vbTab & "PASS"
    Print #fileNo, "Forced Runtime final state" & vbTab & "Runtime" & vbTab & "Inactive" & vbTab & runtimeLeakSource & vbTab & "PASS"
    Print #fileNo, "Scenario Fork forced nesting final state" & vbTab & "Scenario Fork" & vbTab & "Inactive" & vbTab & scenarioForkLeakSource & vbTab & "PASS"
    Close #fileNo

End Sub

'------------------------------------------------------------------------------
' FR: Verifie le contrat WBS Write Guard Harness Reset et signale toute divergence au harnais.
' EN: Verifies the WBS Write Guard Harness Reset contract and reports any divergence to the harness.
'------------------------------------------------------------------------------

Private Sub WBSWriteGuardHarness_Reset()

    Dim safetyCount As Long

    Do While IsAuthorizedWBSWriteActive()
        EndAuthorizedWBSWrite
        safetyCount = safetyCount + 1
        If safetyCount > 32 Then
            Err.Raise vbObjectError + 9611, "WBSWriteGuardHarness", _
                "Guard could not be reset after 32 EndAuthorizedWBSWrite calls."
        End If
    Loop

End Sub

'------------------------------------------------------------------------------
' FR: Verifie le contrat WBS Write Guard Harness Assert et signale toute divergence au harnais.
' EN: Verifies the WBS Write Guard Harness Assert contract and reports any divergence to the harness.
'------------------------------------------------------------------------------

Private Sub WBSWriteGuardHarness_Assert(ByVal condition As Boolean, ByVal messageText As String)

    If Not condition Then
        WBSWriteGuardHarness_Trace "ASSERT FAIL " & messageText
        Err.Raise vbObjectError + 9612, "WBSWriteGuardHarness", messageText
    End If

End Sub

'------------------------------------------------------------------------------
' FR: Verifie le contrat WBS Write Guard Harness Trace et signale toute divergence au harnais.
' EN: Verifies the WBS Write Guard Harness Trace contract and reports any divergence to the harness.
'------------------------------------------------------------------------------

Private Sub WBSWriteGuardHarness_Trace(ByVal messageText As String)

    Dim fileNo As Integer

    If Trim$(gWBSGuardHarnessTracePath) = vbNullString Then Exit Sub

    fileNo = FreeFile
    Open gWBSGuardHarnessTracePath For Append As #fileNo
    Print #fileNo, Format$(Now, "yyyy-mm-dd hh:nn:ss") & " " & messageText
    Close #fileNo

End Sub

'------------------------------------------------------------------------------
' FR: Verifie le contrat WBS Write Guard Harness Delete File et signale toute divergence au harnais.
' EN: Verifies the WBS Write Guard Harness Delete File contract and reports any divergence to the harness.
'------------------------------------------------------------------------------

Private Sub WBSWriteGuardHarness_DeleteFile(ByVal filePath As String)

    On Error Resume Next
    If Len(Dir$(filePath, vbNormal)) > 0 Then Kill filePath
    On Error GoTo 0

End Sub
