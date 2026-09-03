Attribute VB_Name = "mod_RuntimeWorkflowHarness"
Option Explicit

'===============================================================================
' MODULE : mod_RuntimeWorkflowHarness
' DOMAINE / DOMAIN : Validation Harnesses
'
' FR
' Harnais de preuve du contrat Runtime Workflow sur des copies de test.
' N'appartient a aucun workflow produit et ne doit pas etre appele en usage normal.
'
' EN
' Proof harness for the Runtime Workflow contract on test copies.
' Is not production workflow code and must not run during normal use.
'
' CONTRATS / CONTRACTS : RuntimeWorkflowHarness_Smoke
' CALLBACKS EXTERNES / EXTERNAL CALLBACKS : Aucun / None
'===============================================================================


'------------------------------------------------------------------------------
' FR: Execute le smoke test de contrat Runtime Workflow / MacroGuard.
' EN: Runs the Runtime Workflow / MacroGuard contract smoke test.
'------------------------------------------------------------------------------
Public Function RuntimeWorkflowHarness_Smoke() As String

    Dim rootId As String
    Dim childId As String
    Dim rootIdFromChild As String
    Dim messages As Collection
    Dim drained As Collection
    Dim item As Object
    Dim abortRaised As Boolean

    On Error GoTo FailSmoke

    ResetPlanningWorkflowContextSafe "RuntimeWorkflowHarness_Smoke.Start"
    RuntimeWorkflowHarness_Assert Not IsPlanningWorkflowActive(), "workflow inactive after reset"
    RuntimeWorkflowHarness_Assert GetCurrentPlanningWorkflowDepth() = 0, "depth is zero after reset"
    RuntimeWorkflowHarness_Assert GetCurrentPlanningWorkflowId() = "", "current id is empty after reset"

    rootId = BeginPlanningWorkflow("HarnessRoot")
    RuntimeWorkflowHarness_Assert rootId <> "", "root id created"
    RuntimeWorkflowHarness_Assert IsPlanningWorkflowActive(), "workflow active after root begin"
    RuntimeWorkflowHarness_Assert IsRootPlanningWorkflow(), "root workflow is root"
    RuntimeWorkflowHarness_Assert GetCurrentPlanningWorkflowDepth() = 1, "root depth is one"
    RuntimeWorkflowHarness_Assert GetCurrentPlanningRootWorkflowId() = rootId, "root id matches root workflow id"

    childId = BeginPlanningWorkflow("HarnessChild")
    rootIdFromChild = GetCurrentPlanningRootWorkflowId()
    RuntimeWorkflowHarness_Assert childId <> "", "child id created"
    RuntimeWorkflowHarness_Assert childId <> rootId, "child id differs from root id"
    RuntimeWorkflowHarness_Assert Not IsRootPlanningWorkflow(), "child workflow is not root"
    RuntimeWorkflowHarness_Assert GetCurrentPlanningWorkflowDepth() = 2, "child depth is two"
    RuntimeWorkflowHarness_Assert rootIdFromChild = rootId, "child root id points to root"
    RuntimeWorkflowHarness_Assert CanCurrentWorkflowDisplay("RuntimeWorkflowHarness"), "display remains allowed in soft mode"

    EndPlanningWorkflow
    RuntimeWorkflowHarness_Assert GetCurrentPlanningWorkflowDepth() = 1, "depth returns to root after child end"
    EndPlanningWorkflow
    RuntimeWorkflowHarness_Assert Not IsPlanningWorkflowActive(), "workflow inactive after root end"

    BeginPlanningWorkflowStopOnlyDisplay
    RuntimeWorkflowHarness_Assert IsPlanningWorkflowStopOnlyDisplay(), "stop-only display flag set"
    EndPlanningWorkflowStopOnlyDisplay
    RuntimeWorkflowHarness_Assert Not IsPlanningWorkflowStopOnlyDisplay(), "stop-only display flag cleared"

    BeginPlanningWorkflowFinalDisplay
    RuntimeWorkflowHarness_Assert IsPlanningWorkflowFinalDisplay(), "final display flag set"
    EndPlanningWorkflowFinalDisplay
    RuntimeWorkflowHarness_Assert Not IsPlanningWorkflowFinalDisplay(), "final display flag cleared"

    rootId = BeginPlanningWorkflow("Run_Full_Update")
    childId = BeginPlanningWorkflow("HarnessChild")
    RuntimeWorkflowHarness_Assert ShouldDeferCurrentWorkflowDisplayToRoot("RuntimeWorkflowHarness"), "full-update child display defers to root"

    Set messages = New Collection
    Set item = CreateObject("Scripting.Dictionary")
    item("Type") = "STOP"
    item("Message") = "Runtime workflow harness message"
    messages.Add item
    DeferPlanningWorkflowDisplayMessages messages

    Set drained = New Collection
    DrainPlanningWorkflowDeferredDisplayMessages drained
    RuntimeWorkflowHarness_Assert drained.Count = 1, "deferred message drained once"
    Set drained = New Collection
    DrainPlanningWorkflowDeferredDisplayMessages drained
    RuntimeWorkflowHarness_Assert drained.Count = 0, "deferred queue empty after drain"
    EndPlanningWorkflow
    EndPlanningWorkflow

    BeginMacroRun "RuntimeWorkflowHarness"
    RuntimeWorkflowHarness_Assert IsMacroRunActive(), "macro guard active after begin"
    RequestMacroAbort "RuntimeWorkflowHarness", "Arret test", "Test abort"
    RuntimeWorkflowHarness_Assert IsMacroAbortRequested(), "macro abort requested"
    On Error Resume Next
    AbortIfRequested "RuntimeWorkflowHarness.AbortIfRequested"
    abortRaised = (Err.Number <> 0)
    Err.Clear
    On Error GoTo FailSmoke
    RuntimeWorkflowHarness_Assert abortRaised, "AbortIfRequested raises while abort is requested"
    EndMacroRun
    RuntimeWorkflowHarness_Assert Not IsMacroRunActive(), "macro guard inactive after end"
    RuntimeWorkflowHarness_Assert Not IsMacroAbortRequested(), "macro abort cleared after end"

    RuntimeWorkflowHarness_Smoke = "PASS"
    Exit Function

FailSmoke:
    On Error Resume Next
    ResetPlanningWorkflowContextSafe "RuntimeWorkflowHarness_Smoke.Fail"
    If IsMacroRunActive() Then EndMacroRun
    RuntimeWorkflowHarness_Smoke = "FAIL: " & Err.Description
    Err.Raise vbObjectError + 9312, "RuntimeWorkflowHarness_Smoke", RuntimeWorkflowHarness_Smoke

End Function

'------------------------------------------------------------------------------
' FR: Valide une condition du harnais Runtime Workflow.
' EN: Validates one Runtime Workflow harness condition.
'------------------------------------------------------------------------------
Private Sub RuntimeWorkflowHarness_Assert(ByVal condition As Boolean, ByVal messageText As String)

    If Not condition Then
        Err.Raise vbObjectError + 9311, "RuntimeWorkflowHarness_Assert", messageText
    End If

End Sub
