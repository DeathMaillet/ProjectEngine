Attribute VB_Name = "mod_BaselineRexHybridHarness"
Option Explicit

'===============================================================================
' MODULE : mod_BaselineRexHybridHarness
' DOMAIN : REX analytics validation
'
' Synthetic evidence for the hybrid Baseline temporal state contract.
'===============================================================================

Public Function BaselineRexHybridHarness_Export(ByVal evidenceFolder As String) As String

    On Error GoTo Fail

    Dim dataArr As Variant
    Dim mapCalc As Object
    Dim rowById As Object
    Dim predsById As Object
    Dim validIds As Object
    Dim topoOrder As Collection
    Dim predLagBySuccPred As Object
    Dim predTypeBySuccPred As Object
    Dim rexStartById As Object
    Dim rexFinishById As Object
    Dim rexDurationById As Object
    Dim diagnosticsById As Object

    Set mapCalc = CreateObject("Scripting.Dictionary")
    mapCalc("ID") = 1
    mapCalc("WBS") = 2
    mapCalc("Task Type") = 3
    mapCalc("Cal") = 4
    mapCalc("Baseline Start") = 5
    mapCalc("Baseline Finish") = 6
    mapCalc("Baseline Duration") = 7

    ReDim dataArr(1 To 5, 1 To 7)
    BaselineRexHybridHarness_Row dataArr, 1, "A", "1", "Task", "7D", DateSerial(2026, 1, 5), DateSerial(2026, 1, 7), 3
    BaselineRexHybridHarness_Row dataArr, 2, "B", "2", "Task", "7D", Empty, Empty, 2
    BaselineRexHybridHarness_Row dataArr, 3, "C", "3", "Task", "7D", DateSerial(2026, 1, 20), Empty, 4
    BaselineRexHybridHarness_Row dataArr, 4, "D", "4", "Task", "7D", Empty, DateSerial(2026, 1, 31), 5
    BaselineRexHybridHarness_Row dataArr, 5, "E", "5", "Task", "7D", Empty, Empty, Empty

    Set rowById = CreateObject("Scripting.Dictionary")
    Set predsById = CreateObject("Scripting.Dictionary")
    Set validIds = CreateObject("Scripting.Dictionary")
    Set topoOrder = New Collection
    Set predLagBySuccPred = CreateObject("Scripting.Dictionary")
    Set predTypeBySuccPred = CreateObject("Scripting.Dictionary")

    BaselineRexHybridHarness_AddTask rowById, predsById, validIds, topoOrder, "A", 1
    BaselineRexHybridHarness_AddTask rowById, predsById, validIds, topoOrder, "B", 2
    BaselineRexHybridHarness_AddTask rowById, predsById, validIds, topoOrder, "C", 3
    BaselineRexHybridHarness_AddTask rowById, predsById, validIds, topoOrder, "D", 4
    BaselineRexHybridHarness_AddTask rowById, predsById, validIds, topoOrder, "E", 5

    predsById("B").Add "A"
    predLagBySuccPred("B|A") = 0
    predTypeBySuccPred("B|A") = "FS"

    Set rexStartById = CreateObject("Scripting.Dictionary")
    Set rexFinishById = CreateObject("Scripting.Dictionary")
    Set rexDurationById = CreateObject("Scripting.Dictionary")
    Set diagnosticsById = CreateObject("Scripting.Dictionary")

    BuildBaselineRexTemporalState _
        dataArr, mapCalc, rowById, predsById, validIds, topoOrder, _
        predLagBySuccPred, predTypeBySuccPred, _
        rexStartById, rexFinishById, rexDurationById, diagnosticsById

    BaselineRexHybridHarness_WriteMain evidenceFolder, rexStartById, rexFinishById, rexDurationById, diagnosticsById
    BaselineRexHybridHarness_WriteDependencyOnly evidenceFolder, rexStartById, rexFinishById, diagnosticsById
    BaselineRexHybridHarness_WriteExplicitGap evidenceFolder, rexStartById, rexFinishById, diagnosticsById
    BaselineRexHybridHarness_WriteDiagnostics evidenceFolder, diagnosticsById

    BaselineRexHybridHarness_Export = "PASS"
    Exit Function

Fail:
    BaselineRexHybridHarness_Export = "FAIL|" & CStr(Err.Number) & "|" & Err.Description

End Function

Private Sub BaselineRexHybridHarness_Row( _
    ByRef dataArr As Variant, _
    ByVal rowIndex As Long, _
    ByVal idValue As String, _
    ByVal wbsValue As String, _
    ByVal typeValue As String, _
    ByVal calValue As String, _
    ByVal baselineStart As Variant, _
    ByVal baselineFinish As Variant, _
    ByVal baselineDuration As Variant)

    dataArr(rowIndex, 1) = idValue
    dataArr(rowIndex, 2) = wbsValue
    dataArr(rowIndex, 3) = typeValue
    dataArr(rowIndex, 4) = calValue
    dataArr(rowIndex, 5) = baselineStart
    dataArr(rowIndex, 6) = baselineFinish
    dataArr(rowIndex, 7) = baselineDuration

End Sub

Private Sub BaselineRexHybridHarness_AddTask( _
    ByVal rowById As Object, _
    ByVal predsById As Object, _
    ByVal validIds As Object, _
    ByVal topoOrder As Collection, _
    ByVal taskId As String, _
    ByVal rowIndex As Long)

    rowById(taskId) = rowIndex
    Set predsById(taskId) = New Collection
    validIds(taskId) = True
    topoOrder.Add taskId

End Sub

Private Sub BaselineRexHybridHarness_WriteMain( _
    ByVal folderPath As String, _
    ByVal rexStartById As Object, _
    ByVal rexFinishById As Object, _
    ByVal rexDurationById As Object, _
    ByVal diagnosticsById As Object)

    Dim f As Integer
    f = FreeFile
    Open BaselineRexHybridHarness_Path(folderPath, "baseline_hybrid_fixtures.tsv") For Output As #f
    Print #f, "TaskId" & vbTab & "ExpectedSource" & vbTab & "Start" & vbTab & "Finish" & vbTab & "Duration" & vbTab & "Diagnostic" & vbTab & "Status"
    BaselineRexHybridHarness_PrintTask f, "A", "ExplicitStartFinishDuration", rexStartById, rexFinishById, rexDurationById, diagnosticsById, True
    BaselineRexHybridHarness_PrintTask f, "B", "DependencyOnlyFS", rexStartById, rexFinishById, rexDurationById, diagnosticsById, True
    BaselineRexHybridHarness_PrintTask f, "C", "ExplicitStartDuration", rexStartById, rexFinishById, rexDurationById, diagnosticsById, True
    BaselineRexHybridHarness_PrintTask f, "D", "ExplicitFinishDuration", rexStartById, rexFinishById, rexDurationById, diagnosticsById, True
    BaselineRexHybridHarness_PrintTask f, "E", "Incomplete", rexStartById, rexFinishById, rexDurationById, diagnosticsById, False
    Close #f

End Sub

Private Sub BaselineRexHybridHarness_WriteDependencyOnly(ByVal folderPath As String, ByVal rexStartById As Object, ByVal rexFinishById As Object, ByVal diagnosticsById As Object)

    Dim f As Integer
    Dim statusText As String
    f = FreeFile
    statusText = "FAIL"
    If rexStartById.Exists("B") And rexFinishById.Exists("B") And Not diagnosticsById.Exists("B") Then statusText = "PASS"
    Open BaselineRexHybridHarness_Path(folderPath, "rex_dependency_only_tasks.tsv") For Output As #f
    Print #f, "TaskId" & vbTab & "HasExplicitBaselineDates" & vbTab & "HasDuration" & vbTab & "ProjectedStart" & vbTab & "ProjectedFinish" & vbTab & "Status"
    Print #f, "B" & vbTab & "NO" & vbTab & "YES" & vbTab & BaselineRexHybridHarness_Value(rexStartById, "B") & vbTab & BaselineRexHybridHarness_Value(rexFinishById, "B") & vbTab & statusText
    Close #f

End Sub

Private Sub BaselineRexHybridHarness_WriteExplicitGap(ByVal folderPath As String, ByVal rexStartById As Object, ByVal rexFinishById As Object, ByVal diagnosticsById As Object)

    Dim f As Integer
    f = FreeFile
    Open BaselineRexHybridHarness_Path(folderPath, "rex_explicit_gap_preservation.tsv") For Output As #f
    Print #f, "TaskId" & vbTab & "InputShape" & vbTab & "ProjectedStart" & vbTab & "ProjectedFinish" & vbTab & "Status"
    Print #f, "C" & vbTab & "Start+Duration" & vbTab & BaselineRexHybridHarness_Value(rexStartById, "C") & vbTab & BaselineRexHybridHarness_Value(rexFinishById, "C") & vbTab & IIf(diagnosticsById.Exists("C"), "FAIL", "PASS")
    Print #f, "D" & vbTab & "Finish+Duration" & vbTab & BaselineRexHybridHarness_Value(rexStartById, "D") & vbTab & BaselineRexHybridHarness_Value(rexFinishById, "D") & vbTab & IIf(diagnosticsById.Exists("D"), "FAIL", "PASS")
    Close #f

End Sub

Private Sub BaselineRexHybridHarness_WriteDiagnostics(ByVal folderPath As String, ByVal diagnosticsById As Object)

    Dim f As Integer
    f = FreeFile
    Open BaselineRexHybridHarness_Path(folderPath, "rex_input_diagnostics.tsv") For Output As #f
    Print #f, "TaskId" & vbTab & "Diagnostic" & vbTab & "MentionsDurationOnly" & vbTab & "Status"
    Print #f, "E" & vbTab & BaselineRexHybridHarness_Diagnostic(diagnosticsById, "E") & vbTab & "NO" & vbTab & IIf(diagnosticsById.Exists("E"), "PASS", "FAIL")
    Close #f

End Sub

Private Sub BaselineRexHybridHarness_PrintTask(ByVal f As Integer, ByVal taskId As String, ByVal expectedSource As String, ByVal rexStartById As Object, ByVal rexFinishById As Object, ByVal rexDurationById As Object, ByVal diagnosticsById As Object, ByVal shouldProject As Boolean)

    Dim ok As Boolean
    ok = (rexStartById.Exists(taskId) And rexFinishById.Exists(taskId) And rexDurationById.Exists(taskId))
    If Not shouldProject Then ok = diagnosticsById.Exists(taskId)
    Print #f, taskId & vbTab & expectedSource & vbTab & BaselineRexHybridHarness_Value(rexStartById, taskId) & vbTab & BaselineRexHybridHarness_Value(rexFinishById, taskId) & vbTab & BaselineRexHybridHarness_Value(rexDurationById, taskId) & vbTab & BaselineRexHybridHarness_Diagnostic(diagnosticsById, taskId) & vbTab & IIf(ok, "PASS", "FAIL")

End Sub

Private Function BaselineRexHybridHarness_Value(ByVal target As Object, ByVal key As String) As String
    If Not target.Exists(key) Then
        BaselineRexHybridHarness_Value = ""
    ElseIf IsDate(target(key)) Then
        BaselineRexHybridHarness_Value = Format$(CDate(target(key)), "yyyy-mm-dd")
    Else
        BaselineRexHybridHarness_Value = CStr(target(key))
    End If
End Function

Private Function BaselineRexHybridHarness_Diagnostic(ByVal target As Object, ByVal key As String) As String
    If target.Exists(key) Then BaselineRexHybridHarness_Diagnostic = CStr(target(key))
End Function

Private Function BaselineRexHybridHarness_Path(ByVal folderPath As String, ByVal fileName As String) As String
    If Right$(folderPath, 1) = "\" Then
        BaselineRexHybridHarness_Path = folderPath & fileName
    Else
        BaselineRexHybridHarness_Path = folderPath & "\" & fileName
    End If
End Function
