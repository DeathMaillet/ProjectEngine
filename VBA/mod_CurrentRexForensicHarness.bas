Attribute VB_Name = "mod_CurrentRexForensicHarness"
Option Explicit

'===============================================================================
' MODULE : mod_CurrentRexForensicHarness
' DOMAINE / DOMAIN : Analytics forensic harness
'
' FR
' Harnais de preuve pour comparer les analytics Current et REX sans modifier les
' regles produit. Exporte uniquement des TSV depuis une copie d'audit.
'
' EN
' Proof harness comparing Current and REX analytics without changing product
' rules. Exports TSV files from an audit copy only.
'
' CONTRATS / CONTRACTS : CurrentRexForensicHarness_Export
'===============================================================================

Public Function CurrentRexForensicHarness_Export(ByVal evidenceFolder As String) As String

    On Error GoTo ErrHandler

    Dim wsCalc As Worksheet
    Dim tblCalc As ListObject
    Dim mapCalc As Object
    Dim dataArr As Variant
    Dim linksBySuccId As Object
    Dim executionNetwork As clsCompiledExecutionNetwork
    Dim rexStartById As Object
    Dim rexFinishById As Object
    Dim rexDurationById As Object

    If Len(Dir$(evidenceFolder, vbDirectory)) = 0 Then MkDir evidenceFolder

    Set wsCalc = ThisWorkbook.Worksheets("CALC")
    Set tblCalc = wsCalc.ListObjects("tbl_CALC")
    Set mapCalc = CanonicalIdentity_BuildColumnMap(tblCalc)
    dataArr = tblCalc.DataBodyRange.value

    Set linksBySuccId = BuildCoreLinksBySucc_FromLogicLinksTable_Expanded(tblCalc)
    Set executionNetwork = CompileExecutionNetwork(dataArr, mapCalc, linksBySuccId)
    Set rexStartById = CreateObject("Scripting.Dictionary")
    Set rexFinishById = CreateObject("Scripting.Dictionary")
    Set rexDurationById = CreateObject("Scripting.Dictionary")

    CurrentRexForensic_BuildRexForward dataArr, mapCalc, executionNetwork, rexStartById, rexFinishById, rexDurationById

    CurrentRexForensic_WritePipeline evidenceFolder
    CurrentRexForensic_WriteEdges evidenceFolder, dataArr, mapCalc, executionNetwork
    CurrentRexForensic_WriteComponents evidenceFolder, dataArr, mapCalc, executionNetwork
    CurrentRexForensic_WriteTerminals evidenceFolder, dataArr, mapCalc, executionNetwork, rexFinishById
    CurrentRexForensic_WriteDateSources evidenceFolder, dataArr, mapCalc, executionNetwork, rexStartById, rexFinishById
    CurrentRexForensic_WriteRexFormulaTerms evidenceFolder, dataArr, mapCalc, executionNetwork, rexStartById, rexFinishById, rexDurationById, "37"
    CurrentRexForensic_WriteForwardDiff evidenceFolder, dataArr, mapCalc, executionNetwork, rexStartById, rexFinishById
    CurrentRexForensic_WriteFinalDiffs evidenceFolder, dataArr, mapCalc, executionNetwork
    CurrentRexForensic_WriteSynthetic evidenceFolder

    CurrentRexForensicHarness_Export = "PASS"
    Exit Function

ErrHandler:
    CurrentRexForensicHarness_Export = "FAIL|" & CStr(Err.Number) & "|" & Err.Description

End Function

Private Sub CurrentRexForensic_WritePipeline(ByVal folderPath As String)

    Dim f As Integer
    f = FreeFile
    Open CurrentRexForensic_Path(folderPath, "current_vs_rex_pipeline.tsv") For Output As #f
    Print #f, "Stage" & vbTab & "CurrentFunction" & vbTab & "RexFunction" & vbTab & "SameInput" & vbTab & "SameAlgorithm" & vbTab & "SameOutputExpected" & vbTab & "ObservedDifference" & vbTab & "Notes"
    Print #f, "Source dates" & vbTab & "Calculated Start/Finish/Duration" & vbTab & "Baseline Start/Finish/Duration" & vbTab & "NO" & vbTab & "NO" & vbTab & "YES under Baseline-state contract" & vbTab & "YES possible" & vbTab & "REX forward pass consumes the placed Baseline Start/Finish state."
    Print #f, "Logical links" & vbTab & "CompiledExecutionNetwork" & vbTab & "CompiledExecutionNetwork" & vbTab & "YES" & vbTab & "YES" & vbTab & "YES" & vbTab & "NO" & vbTab & "Both pipelines consume same executionNetwork analytics edges."
    Print #f, "Expanded links" & vbTab & "AnalyticsPredsById/ChildrenById" & vbTab & "AnalyticsPredsById/ChildrenById" & vbTab & "YES" & vbTab & "YES" & vbTab & "YES" & vbTab & "NO" & vbTab & "Parent expansion is owned by compiled network."
    Print #f, "Forward pass" & vbTab & "Core calculated dates already materialized" & vbTab & "ComputeCriticalPathREX Baseline-state projection" & vbTab & "NO" & vbTab & "NO" & vbTab & "YES under contract 2" & vbTab & "YES" & vbTab & "No divergence is expected when Current dates equal the placed Baseline state."
    Print #f, "Terminal bounds" & vbTab & "Component max Calculated Finish" & vbTab & "Component max Baseline-state REX Finish" & vbTab & "YES when Current equals Baseline" & vbTab & "SAME projection helper" & vbTab & "YES when Current equals Baseline" & vbTab & "YES possible" & vbTab & "Current and REX use same component policy over different date vectors."
    Print #f, "Backward pass" & vbTab & "Current backward over Calculated dates" & vbTab & "REX backward over placed Baseline-state dates" & vbTab & "YES when Current equals Baseline" & vbTab & "Very similar logic" & vbTab & "YES under contract 2" & vbTab & "YES possible" & vbTab & "Same edges, same ES/EF and terminal anchors under invariant precondition."
    Print #f, "Floats / flags" & vbTab & "TF/FF/CP/LP Current" & vbTab & "TF/FF/CP/LP REX" & vbTab & "NO if dates diverge" & vbTab & "Policy equivalent by family" & vbTab & "YES under contract 2" & vbTab & "YES" & vbTab & "Observed differences should be zero under invariant precondition."
    Close #f

End Sub

Private Sub CurrentRexForensic_WriteEdges( _
    ByVal folderPath As String, _
    ByRef dataArr As Variant, _
    ByVal mapCalc As Object, _
    ByVal net As clsCompiledExecutionNetwork)

    Dim fCur As Integer
    Dim fRex As Integer
    Dim fDiff As Integer
    Dim nodeIds As Variant
    Dim rowsByNode As Variant
    Dim predOffsets As Variant
    Dim predNodes As Variant
    Dim predTypes As Variant
    Dim predLags As Variant
    Dim componentByNode As Variant
    Dim succNode As Long
    Dim edgeIdx As Long
    Dim predNode As Long
    Dim lineText As String

    nodeIds = net.NodeIds
    rowsByNode = net.RowsByNode
    predOffsets = net.PredOffsets
    predNodes = net.PredNodes
    predTypes = net.PredTypes
    predLags = net.PredLags
    componentByNode = net.ComponentByNode

    fCur = FreeFile
    Open CurrentRexForensic_Path(folderPath, "current_edges.tsv") For Output As #fCur
    fRex = FreeFile
    Open CurrentRexForensic_Path(folderPath, "rex_edges.tsv") For Output As #fRex
    Print #fCur, CurrentRexForensic_EdgeHeader()
    Print #fRex, CurrentRexForensic_EdgeHeader()

    For succNode = 1 To net.NodeCount
        For edgeIdx = CLng(predOffsets(succNode)) To CLng(predOffsets(succNode + 1)) - 1
            predNode = CLng(predNodes(edgeIdx))
            lineText = CurrentRexForensic_EdgeLine("SHARED", nodeIds, rowsByNode, componentByNode, predNode, succNode, CStr(predTypes(edgeIdx)), CDbl(predLags(edgeIdx)), dataArr, mapCalc)
            Print #fCur, Replace(lineText, "SHARED", "CURRENT", 1, 1, vbBinaryCompare)
            Print #fRex, Replace(lineText, "SHARED", "REX", 1, 1, vbBinaryCompare)
        Next edgeIdx
    Next succNode

    Close #fCur
    Close #fRex

    fDiff = FreeFile
    Open CurrentRexForensic_Path(folderPath, "edge_diff.tsv") For Output As #fDiff
    Print #fDiff, "Check" & vbTab & "Status" & vbTab & "Notes"
    Print #fDiff, "EdgeSet" & vbTab & "PASS" & vbTab & "Current and REX both consume the same compiled analytics edge arrays in this harness."
    Print #fDiff, "LinkTypeLag" & vbTab & "PASS" & vbTab & "PredTypes and PredLags are exported from the same compiled network for both pipelines."
    Close #fDiff

End Sub

Private Function CurrentRexForensic_EdgeHeader() As String
    CurrentRexForensic_EdgeHeader = "Pipeline" & vbTab & "PredId" & vbTab & "SuccId" & vbTab & "PredWBS" & vbTab & "SuccWBS" & vbTab & "LinkType" & vbTab & "Lag" & vbTab & "Component"
End Function

Private Function CurrentRexForensic_EdgeLine( _
    ByVal pipelineName As String, _
    ByRef nodeIds As Variant, _
    ByRef rowsByNode As Variant, _
    ByRef componentByNode As Variant, _
    ByVal predNode As Long, _
    ByVal succNode As Long, _
    ByVal linkType As String, _
    ByVal lagValue As Double, _
    ByRef dataArr As Variant, _
    ByVal mapCalc As Object) As String

    CurrentRexForensic_EdgeLine = _
        pipelineName & vbTab & _
        CurrentRexForensic_Tsv(CStr(nodeIds(predNode))) & vbTab & _
        CurrentRexForensic_Tsv(CStr(nodeIds(succNode))) & vbTab & _
        CurrentRexForensic_Tsv(CStr(dataArr(rowsByNode(predNode), mapCalc("WBS")))) & vbTab & _
        CurrentRexForensic_Tsv(CStr(dataArr(rowsByNode(succNode), mapCalc("WBS")))) & vbTab & _
        CurrentRexForensic_Tsv(linkType) & vbTab & _
        CStr(lagValue) & vbTab & _
        CStr(componentByNode(succNode))

End Function

Private Sub CurrentRexForensic_WriteComponents( _
    ByVal folderPath As String, _
    ByRef dataArr As Variant, _
    ByVal mapCalc As Object, _
    ByVal net As clsCompiledExecutionNetwork)

    Dim fCur As Integer
    Dim fRex As Integer
    Dim fDiff As Integer
    Dim nodeIds As Variant
    Dim rowsByNode As Variant
    Dim componentByNode As Variant
    Dim nodeIndex As Long
    Dim lineText As String

    nodeIds = net.NodeIds
    rowsByNode = net.RowsByNode
    componentByNode = net.ComponentByNode

    fCur = FreeFile
    Open CurrentRexForensic_Path(folderPath, "current_components.tsv") For Output As #fCur
    fRex = FreeFile
    Open CurrentRexForensic_Path(folderPath, "rex_components.tsv") For Output As #fRex
    Print #fCur, "Pipeline" & vbTab & "ID" & vbTab & "WBS" & vbTab & "Component"
    Print #fRex, "Pipeline" & vbTab & "ID" & vbTab & "WBS" & vbTab & "Component"
    For nodeIndex = 1 To net.NodeCount
        lineText = "{PIPE}" & vbTab & CurrentRexForensic_Tsv(CStr(nodeIds(nodeIndex))) & vbTab & CurrentRexForensic_Tsv(CStr(dataArr(rowsByNode(nodeIndex), mapCalc("WBS")))) & vbTab & CStr(componentByNode(nodeIndex))
        Print #fCur, Replace(lineText, "{PIPE}", "CURRENT")
        Print #fRex, Replace(lineText, "{PIPE}", "REX")
    Next nodeIndex
    Close #fCur
    Close #fRex

    fDiff = FreeFile
    Open CurrentRexForensic_Path(folderPath, "component_diff.tsv") For Output As #fDiff
    Print #fDiff, "Check" & vbTab & "Status" & vbTab & "Notes"
    Print #fDiff, "ComponentAssignment" & vbTab & "PASS" & vbTab & "Current and REX use the same ComponentByNode array."
    Close #fDiff

End Sub

Private Sub CurrentRexForensic_WriteTerminals( _
    ByVal folderPath As String, _
    ByRef dataArr As Variant, _
    ByVal mapCalc As Object, _
    ByVal net As clsCompiledExecutionNetwork, _
    ByVal rexFinishById As Object)

    Dim currentFinish As Object
    Dim rexFinish As Object
    Dim fCur As Integer
    Dim fRex As Integer
    Dim fDiff As Integer
    Dim idKey As Variant
    Dim currentVal As Variant
    Dim rexVal As Variant

    Set currentFinish = CompiledNetwork_BuildCurrentFinishById(net, dataArr, mapCalc)
    Set rexFinish = CompiledNetwork_BuildFinishByIdFromValues(net, rexFinishById)

    fCur = FreeFile
    Open CurrentRexForensic_Path(folderPath, "current_terminals.tsv") For Output As #fCur
    fRex = FreeFile
    Open CurrentRexForensic_Path(folderPath, "rex_terminals.tsv") For Output As #fRex
    Print #fCur, "Pipeline" & vbTab & "ID" & vbTab & "TerminalAnchor"
    Print #fRex, "Pipeline" & vbTab & "ID" & vbTab & "TerminalAnchor"
    For Each idKey In net.RowById.Keys
        currentVal = Empty
        If currentFinish.Exists(CStr(idKey)) Then currentVal = currentFinish(CStr(idKey))
        rexVal = Empty
        If rexFinish.Exists(CStr(idKey)) Then rexVal = rexFinish(CStr(idKey))
        Print #fCur, "CURRENT" & vbTab & CurrentRexForensic_Tsv(CStr(idKey)) & vbTab & CurrentRexForensic_Tsv(CurrentRexForensic_ValueText(currentVal))
        Print #fRex, "REX" & vbTab & CurrentRexForensic_Tsv(CStr(idKey)) & vbTab & CurrentRexForensic_Tsv(CurrentRexForensic_ValueText(rexVal))
    Next idKey
    Close #fCur
    Close #fRex

    fDiff = FreeFile
    Open CurrentRexForensic_Path(folderPath, "terminal_diff.tsv") For Output As #fDiff
    Print #fDiff, "ID" & vbTab & "CurrentAnchor" & vbTab & "RexAnchor" & vbTab & "Status"
    For Each idKey In net.RowById.Keys
        currentVal = Empty
        rexVal = Empty
        If currentFinish.Exists(CStr(idKey)) Then currentVal = currentFinish(CStr(idKey))
        If rexFinish.Exists(CStr(idKey)) Then rexVal = rexFinish(CStr(idKey))
        If CurrentRexForensic_SameValue(currentVal, rexVal) Then
            Print #fDiff, CurrentRexForensic_Tsv(CStr(idKey)) & vbTab & CurrentRexForensic_Tsv(CurrentRexForensic_ValueText(currentVal)) & vbTab & CurrentRexForensic_Tsv(CurrentRexForensic_ValueText(rexVal)) & vbTab & "SAME"
        Else
            Print #fDiff, CurrentRexForensic_Tsv(CStr(idKey)) & vbTab & CurrentRexForensic_Tsv(CurrentRexForensic_ValueText(currentVal)) & vbTab & CurrentRexForensic_Tsv(CurrentRexForensic_ValueText(rexVal)) & vbTab & "DIFF"
        End If
    Next idKey
    Close #fDiff

End Sub

Private Sub CurrentRexForensic_BuildRexForward( _
    ByRef dataArr As Variant, _
    ByVal mapCalc As Object, _
    ByVal net As clsCompiledExecutionNetwork, _
    ByVal rexStartById As Object, _
    ByVal rexFinishById As Object, _
    ByVal rexDurationById As Object)

    Dim diagnosticsById As Object

    Set diagnosticsById = CreateObject("Scripting.Dictionary")

    BuildBaselineRexTemporalState _
        dataArr, mapCalc, net.RowById, net.AnalyticsPredsById, net.ValidLeafIds, net.TopoOrder, _
        net.PredLagBySuccPred, net.PredTypeBySuccPred, _
        rexStartById, rexFinishById, rexDurationById, diagnosticsById

End Sub
Private Sub CurrentRexForensic_WriteDateSources( _
    ByVal folderPath As String, _
    ByRef dataArr As Variant, _
    ByVal mapCalc As Object, _
    ByVal net As clsCompiledExecutionNetwork, _
    ByVal rexStartById As Object, _
    ByVal rexFinishById As Object)

    Dim f As Integer
    Dim idKey As Variant
    Dim rowIndex As Long
    Dim taskId As String

    f = FreeFile
    Open CurrentRexForensic_Path(folderPath, "rex_date_sources.tsv") For Output As #f
    Print #f, "ID" & vbTab & "WBS" & vbTab & "BaselineStart" & vbTab & "BaselineFinish" & vbTab & "BaselineDuration" & vbTab & "CalculatedStart" & vbTab & "CalculatedFinish" & vbTab & "CalculatedDuration" & vbTab & "RexStartUsed" & vbTab & "RexFinishUsed"
    For Each idKey In net.RowById.Keys
        taskId = CStr(idKey)
        rowIndex = CLng(net.RowById(taskId))
        Print #f, CurrentRexForensic_Tsv(taskId) & vbTab & _
            CurrentRexForensic_Tsv(CStr(dataArr(rowIndex, mapCalc("WBS")))) & vbTab & _
            CurrentRexForensic_Tsv(CurrentRexForensic_Cell(dataArr, rowIndex, mapCalc, "Baseline Start")) & vbTab & _
            CurrentRexForensic_Tsv(CurrentRexForensic_Cell(dataArr, rowIndex, mapCalc, "Baseline Finish")) & vbTab & _
            CurrentRexForensic_Tsv(CurrentRexForensic_Cell(dataArr, rowIndex, mapCalc, "Baseline Duration")) & vbTab & _
            CurrentRexForensic_Tsv(CurrentRexForensic_Cell(dataArr, rowIndex, mapCalc, "Calculated Start")) & vbTab & _
            CurrentRexForensic_Tsv(CurrentRexForensic_Cell(dataArr, rowIndex, mapCalc, "Calculated Finish")) & vbTab & _
            CurrentRexForensic_Tsv(CurrentRexForensic_Cell(dataArr, rowIndex, mapCalc, "Calculated Duration")) & vbTab & _
            CurrentRexForensic_Tsv(CurrentRexForensic_DictValue(rexStartById, taskId)) & vbTab & _
            CurrentRexForensic_Tsv(CurrentRexForensic_DictValue(rexFinishById, taskId))
    Next idKey
    Close #f

End Sub

Private Sub CurrentRexForensic_WriteRexFormulaTerms( _
    ByVal folderPath As String, _
    ByRef dataArr As Variant, _
    ByVal mapCalc As Object, _
    ByVal net As clsCompiledExecutionNetwork, _
    ByVal rexStartById As Object, _
    ByVal rexFinishById As Object, _
    ByVal rexDurationById As Object, _
    ByVal targetId As String)

    Dim f As Integer
    Dim reverseTopo As Collection
    Dim networkFinishById As Object
    Dim rexLateStartById As Object
    Dim rexLateFinishById As Object
    Dim idKey As Variant
    Dim succId As Variant
    Dim taskId As String
    Dim rowIndex As Long
    Dim succRowIndex As Long
    Dim baselineDuration As Variant
    Dim taskCal As String
    Dim succCal As String
    Dim candidateLateFinish As Variant
    Dim candidateLateStart As Variant
    Dim candidateDate As Variant
    Dim effectiveLag As Double
    Dim linkType As String
    Dim linkKey As String
    Dim r As Long
    Dim tfRaw As Variant
    Dim tfValue As Variant
    Dim requiredDate As Variant
    Dim candidateFF As Variant

    Set reverseTopo = New Collection
    Set networkFinishById = CompiledNetwork_BuildFinishByIdFromValues(net, rexFinishById)
    Set rexLateStartById = CreateObject("Scripting.Dictionary")
    Set rexLateFinishById = CreateObject("Scripting.Dictionary")

    For r = net.TopoOrder.Count To 1 Step -1
        reverseTopo.Add CStr(net.TopoOrder(r))
    Next r

    For Each idKey In reverseTopo
        taskId = CStr(idKey)
        If Not net.RowById.Exists(taskId) Then GoTo NextBackward
        rowIndex = CLng(net.RowById(taskId))
        If rexDurationById.Exists(taskId) Then
            baselineDuration = rexDurationById(taskId)
        Else
            baselineDuration = GetCellValue(dataArr(rowIndex, mapCalc("Baseline Duration")))
        End If
        If Not HasValue(baselineDuration) Then GoTo NextBackward

        taskCal = NormalizeCalendarType(dataArr(rowIndex, mapCalc("Cal")))
        candidateLateFinish = Empty

        For Each succId In net.AnalyticsChildrenById(taskId)
            If net.ValidLeafIds.Exists(CStr(succId)) Then
                linkKey = CStr(succId) & "|" & taskId
                If net.PredLagBySuccPred.Exists(linkKey) And net.PredTypeBySuccPred.Exists(linkKey) Then
                    effectiveLag = CDbl(net.PredLagBySuccPred(linkKey))
                    linkType = CStr(net.PredTypeBySuccPred(linkKey))
                    succCal = NormalizeCalendarType(dataArr(CLng(net.RowById(CStr(succId))), mapCalc("Cal")))

                    Select Case linkType
                        Case "SS"
                            If rexLateStartById.Exists(CStr(succId)) Then
                                candidateLateStart = OffsetWorkingDays(rexLateStartById(CStr(succId)), -effectiveLag, succCal)
                                If Not HasValue(candidateLateFinish) Then
                                    candidateLateFinish = AddWorkingDays(candidateLateStart, baselineDuration, taskCal)
                                ElseIf CDbl(AddWorkingDays(candidateLateStart, baselineDuration, taskCal)) < CDbl(candidateLateFinish) Then
                                    candidateLateFinish = AddWorkingDays(candidateLateStart, baselineDuration, taskCal)
                                End If
                            End If

                        Case "FF"
                            If rexLateFinishById.Exists(CStr(succId)) Then
                                If Not HasValue(candidateLateFinish) Then
                                    candidateLateFinish = OffsetWorkingDays(rexLateFinishById(CStr(succId)), -effectiveLag, succCal)
                                ElseIf CDbl(OffsetWorkingDays(rexLateFinishById(CStr(succId)), -effectiveLag, succCal)) < CDbl(candidateLateFinish) Then
                                    candidateLateFinish = OffsetWorkingDays(rexLateFinishById(CStr(succId)), -effectiveLag, succCal)
                                End If
                            End If

                        Case Else
                            If rexLateStartById.Exists(CStr(succId)) Then
                                If Not HasValue(candidateLateFinish) Then
                                    candidateDate = OffsetWorkingDays(rexLateStartById(CStr(succId)), -effectiveLag, succCal)
                                    candidateLateFinish = PreviousWorkingDay(candidateDate, succCal)
                                ElseIf CDbl(PreviousWorkingDay(OffsetWorkingDays(rexLateStartById(CStr(succId)), -effectiveLag, succCal), succCal)) < CDbl(candidateLateFinish) Then
                                    candidateDate = OffsetWorkingDays(rexLateStartById(CStr(succId)), -effectiveLag, succCal)
                                    candidateLateFinish = PreviousWorkingDay(candidateDate, succCal)
                                End If
                            End If
                    End Select
                End If
            End If
        Next succId

        If Not HasValue(candidateLateFinish) Then
            If networkFinishById.Exists(taskId) Then
                rexLateFinishById(taskId) = networkFinishById(taskId)
            End If
        Else
            rexLateFinishById(taskId) = candidateLateFinish
        End If

        If rexLateFinishById.Exists(taskId) Then
            rexLateStartById(taskId) = SubtractWorkingDays(rexLateFinishById(taskId), baselineDuration, taskCal)
        End If

NextBackward:
    Next idKey

    f = FreeFile
    Open CurrentRexForensic_Path(folderPath, "rex_target_formula_terms.tsv") For Output As #f
    Print #f, "Role" & vbTab & "ID" & vbTab & "WBS" & vbTab & "Calendar" & vbTab & "BaselineStart" & vbTab & "BaselineFinish" & vbTab & "BaselineDuration" & vbTab & "RexES" & vbTab & "RexEF" & vbTab & "RexLS" & vbTab & "RexLF" & vbTab & "TF_REX_Formula" & vbTab & "TF_REX_Result" & vbTab & "SuccID" & vbTab & "SuccWBS" & vbTab & "LinkType" & vbTab & "Lag" & vbTab & "RequiredDate" & vbTab & "SuccRexES" & vbTab & "SuccRexEF" & vbTab & "FF_REX_Formula" & vbTab & "FF_REX_Result"

    If net.RowById.Exists(targetId) Then
        rowIndex = CLng(net.RowById(targetId))
        taskCal = NormalizeCalendarType(dataArr(rowIndex, mapCalc("Cal")))
        tfRaw = Empty
        tfValue = Empty
        If rexStartById.Exists(targetId) And rexLateStartById.Exists(targetId) Then
            tfRaw = SignedWorkingDayDelta(rexStartById(targetId), rexLateStartById(targetId), taskCal)
            tfValue = CDbl(tfRaw)
        End If

        Print #f, "TARGET_TOTAL_FLOAT" & vbTab & _
            CurrentRexForensic_Tsv(targetId) & vbTab & _
            CurrentRexForensic_Tsv(CStr(dataArr(rowIndex, mapCalc("WBS")))) & vbTab & _
            CurrentRexForensic_Tsv(taskCal) & vbTab & _
            CurrentRexForensic_Tsv(CurrentRexForensic_Cell(dataArr, rowIndex, mapCalc, "Baseline Start")) & vbTab & _
            CurrentRexForensic_Tsv(CurrentRexForensic_Cell(dataArr, rowIndex, mapCalc, "Baseline Finish")) & vbTab & _
            CurrentRexForensic_Tsv(CurrentRexForensic_Cell(dataArr, rowIndex, mapCalc, "Baseline Duration")) & vbTab & _
            CurrentRexForensic_Tsv(CurrentRexForensic_DictValue(rexStartById, targetId)) & vbTab & _
            CurrentRexForensic_Tsv(CurrentRexForensic_DictValue(rexFinishById, targetId)) & vbTab & _
            CurrentRexForensic_Tsv(CurrentRexForensic_DictValue(rexLateStartById, targetId)) & vbTab & _
            CurrentRexForensic_Tsv(CurrentRexForensic_DictValue(rexLateFinishById, targetId)) & vbTab & _
            CurrentRexForensic_Tsv("SignedWorkingDayDelta(RexES,RexLS,Cal)") & vbTab & _
            CurrentRexForensic_Tsv(CurrentRexForensic_ValueText(tfValue)) & vbTab & _
            "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & ""

        For Each succId In net.AnalyticsChildrenById(targetId)
            If net.ValidLeafIds.Exists(CStr(succId)) Then
                linkKey = CStr(succId) & "|" & targetId
                If net.PredLagBySuccPred.Exists(linkKey) And net.PredTypeBySuccPred.Exists(linkKey) Then
                    succRowIndex = CLng(net.RowById(CStr(succId)))
                    succCal = NormalizeCalendarType(dataArr(succRowIndex, mapCalc("Cal")))
                    effectiveLag = CDbl(net.PredLagBySuccPred(linkKey))
                    linkType = CStr(net.PredTypeBySuccPred(linkKey))
                    Select Case linkType
                        Case "SS"
                            requiredDate = ApplyLag(rexStartById(targetId), effectiveLag, succCal, "SS")
                            candidateFF = SignedWorkingDayOffset(requiredDate, rexStartById(CStr(succId)), succCal)
                        Case "FF"
                            requiredDate = ApplyLag(rexFinishById(targetId), effectiveLag, succCal, "FF")
                            candidateFF = SignedWorkingDayOffset(requiredDate, rexFinishById(CStr(succId)), succCal)
                        Case Else
                            requiredDate = ApplyLag(rexFinishById(targetId), effectiveLag, succCal, "FS")
                            candidateFF = SignedWorkingDayOffset(requiredDate, rexStartById(CStr(succId)), succCal)
                    End Select

                    Print #f, "TARGET_FREE_FLOAT_CANDIDATE" & vbTab & _
                        CurrentRexForensic_Tsv(targetId) & vbTab & _
                        CurrentRexForensic_Tsv(CStr(dataArr(rowIndex, mapCalc("WBS")))) & vbTab & _
                        CurrentRexForensic_Tsv(taskCal) & vbTab & _
                        CurrentRexForensic_Tsv(CurrentRexForensic_Cell(dataArr, rowIndex, mapCalc, "Baseline Start")) & vbTab & _
                        CurrentRexForensic_Tsv(CurrentRexForensic_Cell(dataArr, rowIndex, mapCalc, "Baseline Finish")) & vbTab & _
                        CurrentRexForensic_Tsv(CurrentRexForensic_Cell(dataArr, rowIndex, mapCalc, "Baseline Duration")) & vbTab & _
                        CurrentRexForensic_Tsv(CurrentRexForensic_DictValue(rexStartById, targetId)) & vbTab & _
                        CurrentRexForensic_Tsv(CurrentRexForensic_DictValue(rexFinishById, targetId)) & vbTab & _
                        CurrentRexForensic_Tsv(CurrentRexForensic_DictValue(rexLateStartById, targetId)) & vbTab & _
                        CurrentRexForensic_Tsv(CurrentRexForensic_DictValue(rexLateFinishById, targetId)) & vbTab & _
                        "" & vbTab & "" & vbTab & _
                        CurrentRexForensic_Tsv(CStr(succId)) & vbTab & _
                        CurrentRexForensic_Tsv(CStr(dataArr(succRowIndex, mapCalc("WBS")))) & vbTab & _
                        CurrentRexForensic_Tsv(linkType) & vbTab & _
                        CStr(effectiveLag) & vbTab & _
                        CurrentRexForensic_Tsv(CurrentRexForensic_ValueText(requiredDate)) & vbTab & _
                        CurrentRexForensic_Tsv(CurrentRexForensic_DictValue(rexStartById, CStr(succId))) & vbTab & _
                        CurrentRexForensic_Tsv(CurrentRexForensic_DictValue(rexFinishById, CStr(succId))) & vbTab & _
                        CurrentRexForensic_Tsv("SignedWorkingDayOffset(RequiredDate,SuccRexDate,SuccCal)") & vbTab & _
                        CurrentRexForensic_Tsv(CurrentRexForensic_ValueText(candidateFF))
                End If
            End If
        Next succId
    End If

    Close #f

End Sub

Private Sub CurrentRexForensic_WriteForwardDiff( _
    ByVal folderPath As String, _
    ByRef dataArr As Variant, _
    ByVal mapCalc As Object, _
    ByVal net As clsCompiledExecutionNetwork, _
    ByVal rexStartById As Object, _
    ByVal rexFinishById As Object)

    Dim fCur As Integer
    Dim fRex As Integer
    Dim fDiff As Integer
    Dim idKey As Variant
    Dim rowIndex As Long
    Dim taskId As String
    Dim calcStart As Variant
    Dim calcFinish As Variant
    Dim rexStart As Variant
    Dim rexFinish As Variant
    Dim status As String

    fCur = FreeFile
    Open CurrentRexForensic_Path(folderPath, "forward_pass_current.tsv") For Output As #fCur
    fRex = FreeFile
    Open CurrentRexForensic_Path(folderPath, "forward_pass_rex.tsv") For Output As #fRex
    fDiff = FreeFile
    Open CurrentRexForensic_Path(folderPath, "forward_pass_diff.tsv") For Output As #fDiff
    Print #fCur, "ID" & vbTab & "WBS" & vbTab & "ES" & vbTab & "EF"
    Print #fRex, "ID" & vbTab & "WBS" & vbTab & "ES" & vbTab & "EF"
    Print #fDiff, "ID" & vbTab & "WBS" & vbTab & "CurrentES" & vbTab & "RexES" & vbTab & "CurrentEF" & vbTab & "RexEF" & vbTab & "Status"

    For Each idKey In net.TopoOrder
        taskId = CStr(idKey)
        rowIndex = CLng(net.RowById(taskId))
        calcStart = GetCellValue(dataArr(rowIndex, mapCalc("Calculated Start")))
        calcFinish = GetCellValue(dataArr(rowIndex, mapCalc("Calculated Finish")))
        rexStart = Empty
        rexFinish = Empty
        If rexStartById.Exists(taskId) Then rexStart = rexStartById(taskId)
        If rexFinishById.Exists(taskId) Then rexFinish = rexFinishById(taskId)
        Print #fCur, CurrentRexForensic_Tsv(taskId) & vbTab & CurrentRexForensic_Tsv(CStr(dataArr(rowIndex, mapCalc("WBS")))) & vbTab & CurrentRexForensic_Tsv(CurrentRexForensic_ValueText(calcStart)) & vbTab & CurrentRexForensic_Tsv(CurrentRexForensic_ValueText(calcFinish))
        Print #fRex, CurrentRexForensic_Tsv(taskId) & vbTab & CurrentRexForensic_Tsv(CStr(dataArr(rowIndex, mapCalc("WBS")))) & vbTab & CurrentRexForensic_Tsv(CurrentRexForensic_ValueText(rexStart)) & vbTab & CurrentRexForensic_Tsv(CurrentRexForensic_ValueText(rexFinish))
        If CurrentRexForensic_SameValue(calcStart, rexStart) And CurrentRexForensic_SameValue(calcFinish, rexFinish) Then status = "SAME" Else status = "DIFF"
        Print #fDiff, CurrentRexForensic_Tsv(taskId) & vbTab & CurrentRexForensic_Tsv(CStr(dataArr(rowIndex, mapCalc("WBS")))) & vbTab & CurrentRexForensic_Tsv(CurrentRexForensic_ValueText(calcStart)) & vbTab & CurrentRexForensic_Tsv(CurrentRexForensic_ValueText(rexStart)) & vbTab & CurrentRexForensic_Tsv(CurrentRexForensic_ValueText(calcFinish)) & vbTab & CurrentRexForensic_Tsv(CurrentRexForensic_ValueText(rexFinish)) & vbTab & status
    Next idKey

    Close #fCur
    Close #fRex
    Close #fDiff

End Sub

Private Sub CurrentRexForensic_WriteFinalDiffs( _
    ByVal folderPath As String, _
    ByRef dataArr As Variant, _
    ByVal mapCalc As Object, _
    ByVal net As clsCompiledExecutionNetwork)

    CurrentRexForensic_WriteInvariant folderPath, dataArr, mapCalc, net
    CurrentRexForensic_WriteBackward folderPath, dataArr, mapCalc, net
    CurrentRexForensic_WriteColumnDiff folderPath, dataArr, mapCalc, net, "Critical Path", "Critical Path REX", "critical_path_diff.tsv"
    CurrentRexForensic_WriteColumnDiff folderPath, dataArr, mapCalc, net, "Free Float", "Free Float REX", "free_float_diff.tsv"
    If mapCalc.Exists("Longest Path REX") Then
        CurrentRexForensic_WriteColumnDiff folderPath, dataArr, mapCalc, net, "Longest Path", "Longest Path REX", "longest_path_diff.tsv"
    Else
        CurrentRexForensic_WriteMissingColumn folderPath, "longest_path_diff.tsv", "Longest Path REX column is absent from tbl_CALC in this workbook/schema."
    End If
    CurrentRexForensic_WriteSummaryRollup folderPath, dataArr, mapCalc

End Sub

Private Sub CurrentRexForensic_WriteInvariant( _
    ByVal folderPath As String, _
    ByRef dataArr As Variant, _
    ByVal mapCalc As Object, _
    ByVal net As clsCompiledExecutionNetwork)

    Dim f As Integer
    Dim idKey As Variant
    Dim rowIndex As Long
    Dim taskId As String
    Dim eligible As Boolean
    Dim violation As Boolean
    Dim status As String
    Dim lpRexText As String

    f = FreeFile
    Open CurrentRexForensic_Path(folderPath, "invariant_results.tsv") For Output As #f
    Print #f, "ID" & vbTab & "WBS" & vbTab & "EligibleNoActualForecastBaselineEqualsCalculated" & vbTab & "TF" & vbTab & "TF_REX" & vbTab & "FF" & vbTab & "FF_REX" & vbTab & "CP" & vbTab & "CP_REX" & vbTab & "LP" & vbTab & "LP_REX" & vbTab & "Status"
    For Each idKey In net.ValidLeafIds.Keys
        taskId = CStr(idKey)
        rowIndex = CLng(net.RowById(taskId))
        eligible = CurrentRexForensic_InvariantEligible(dataArr, mapCalc, rowIndex)
        violation = False
        If eligible Then
            If Not CurrentRexForensic_SameCell(dataArr, rowIndex, mapCalc, "Total Float", "Total Float REX") Then violation = True
            If Not CurrentRexForensic_SameCell(dataArr, rowIndex, mapCalc, "Free Float", "Free Float REX") Then violation = True
            If Not CurrentRexForensic_SameCell(dataArr, rowIndex, mapCalc, "Critical Path", "Critical Path REX") Then violation = True
            If mapCalc.Exists("Longest Path REX") Then
                If Not CurrentRexForensic_SameCell(dataArr, rowIndex, mapCalc, "Longest Path", "Longest Path REX") Then violation = True
            End If
        End If
        If Not eligible Then
            status = "NOT_ELIGIBLE"
        ElseIf violation Then
            status = "VIOLATION_UNDER_CONTRACT_2"
        Else
            status = "PASS"
        End If
        lpRexText = ""
        If mapCalc.Exists("Longest Path REX") Then lpRexText = CurrentRexForensic_Cell(dataArr, rowIndex, mapCalc, "Longest Path REX")

        Print #f, CurrentRexForensic_Tsv(taskId) & vbTab & CurrentRexForensic_Tsv(CStr(dataArr(rowIndex, mapCalc("WBS")))) & vbTab & CStr(eligible) & vbTab & _
            CurrentRexForensic_Tsv(CurrentRexForensic_Cell(dataArr, rowIndex, mapCalc, "Total Float")) & vbTab & _
            CurrentRexForensic_Tsv(CurrentRexForensic_Cell(dataArr, rowIndex, mapCalc, "Total Float REX")) & vbTab & _
            CurrentRexForensic_Tsv(CurrentRexForensic_Cell(dataArr, rowIndex, mapCalc, "Free Float")) & vbTab & _
            CurrentRexForensic_Tsv(CurrentRexForensic_Cell(dataArr, rowIndex, mapCalc, "Free Float REX")) & vbTab & _
            CurrentRexForensic_Tsv(CurrentRexForensic_Cell(dataArr, rowIndex, mapCalc, "Critical Path")) & vbTab & _
            CurrentRexForensic_Tsv(CurrentRexForensic_Cell(dataArr, rowIndex, mapCalc, "Critical Path REX")) & vbTab & _
            CurrentRexForensic_Tsv(CurrentRexForensic_Cell(dataArr, rowIndex, mapCalc, "Longest Path")) & vbTab & _
            CurrentRexForensic_Tsv(lpRexText) & vbTab & status
    Next idKey
    Close #f

End Sub

Private Sub CurrentRexForensic_WriteBackward( _
    ByVal folderPath As String, _
    ByRef dataArr As Variant, _
    ByVal mapCalc As Object, _
    ByVal net As clsCompiledExecutionNetwork)

    Dim fCur As Integer
    Dim fRex As Integer
    Dim fDiff As Integer
    Dim idKey As Variant
    Dim rowIndex As Long
    Dim taskId As String
    Dim status As String

    fCur = FreeFile
    Open CurrentRexForensic_Path(folderPath, "backward_pass_current.tsv") For Output As #fCur
    fRex = FreeFile
    Open CurrentRexForensic_Path(folderPath, "backward_pass_rex.tsv") For Output As #fRex
    fDiff = FreeFile
    Open CurrentRexForensic_Path(folderPath, "backward_pass_diff.tsv") For Output As #fDiff
    Print #fCur, "ID" & vbTab & "WBS" & vbTab & "TF" & vbTab & "FF"
    Print #fRex, "ID" & vbTab & "WBS" & vbTab & "TF" & vbTab & "FF"
    Print #fDiff, "ID" & vbTab & "WBS" & vbTab & "CurrentTF" & vbTab & "RexTF" & vbTab & "CurrentFF" & vbTab & "RexFF" & vbTab & "Status"
    For Each idKey In net.ValidLeafIds.Keys
        taskId = CStr(idKey)
        rowIndex = CLng(net.RowById(taskId))
        Print #fCur, CurrentRexForensic_Tsv(taskId) & vbTab & CurrentRexForensic_Tsv(CStr(dataArr(rowIndex, mapCalc("WBS")))) & vbTab & CurrentRexForensic_Tsv(CurrentRexForensic_Cell(dataArr, rowIndex, mapCalc, "Total Float")) & vbTab & CurrentRexForensic_Tsv(CurrentRexForensic_Cell(dataArr, rowIndex, mapCalc, "Free Float"))
        Print #fRex, CurrentRexForensic_Tsv(taskId) & vbTab & CurrentRexForensic_Tsv(CStr(dataArr(rowIndex, mapCalc("WBS")))) & vbTab & CurrentRexForensic_Tsv(CurrentRexForensic_Cell(dataArr, rowIndex, mapCalc, "Total Float REX")) & vbTab & CurrentRexForensic_Tsv(CurrentRexForensic_Cell(dataArr, rowIndex, mapCalc, "Free Float REX"))
        If CurrentRexForensic_SameCell(dataArr, rowIndex, mapCalc, "Total Float", "Total Float REX") And CurrentRexForensic_SameCell(dataArr, rowIndex, mapCalc, "Free Float", "Free Float REX") Then status = "SAME" Else status = "DIFF"
        Print #fDiff, CurrentRexForensic_Tsv(taskId) & vbTab & CurrentRexForensic_Tsv(CStr(dataArr(rowIndex, mapCalc("WBS")))) & vbTab & CurrentRexForensic_Tsv(CurrentRexForensic_Cell(dataArr, rowIndex, mapCalc, "Total Float")) & vbTab & CurrentRexForensic_Tsv(CurrentRexForensic_Cell(dataArr, rowIndex, mapCalc, "Total Float REX")) & vbTab & CurrentRexForensic_Tsv(CurrentRexForensic_Cell(dataArr, rowIndex, mapCalc, "Free Float")) & vbTab & CurrentRexForensic_Tsv(CurrentRexForensic_Cell(dataArr, rowIndex, mapCalc, "Free Float REX")) & vbTab & status
    Next idKey
    Close #fCur
    Close #fRex
    Close #fDiff

End Sub

Private Sub CurrentRexForensic_WriteColumnDiff( _
    ByVal folderPath As String, _
    ByRef dataArr As Variant, _
    ByVal mapCalc As Object, _
    ByVal net As clsCompiledExecutionNetwork, _
    ByVal currentColumn As String, _
    ByVal rexColumn As String, _
    ByVal fileName As String)

    Dim f As Integer
    Dim idKey As Variant
    Dim rowIndex As Long
    Dim status As String

    f = FreeFile
    Open CurrentRexForensic_Path(folderPath, fileName) For Output As #f
    Print #f, "ID" & vbTab & "WBS" & vbTab & "Current" & vbTab & "REX" & vbTab & "Status"
    For Each idKey In net.ValidLeafIds.Keys
        rowIndex = CLng(net.RowById(CStr(idKey)))
        If CurrentRexForensic_SameCell(dataArr, rowIndex, mapCalc, currentColumn, rexColumn) Then status = "SAME" Else status = "DIFF"
        Print #f, CurrentRexForensic_Tsv(CStr(idKey)) & vbTab & CurrentRexForensic_Tsv(CStr(dataArr(rowIndex, mapCalc("WBS")))) & vbTab & CurrentRexForensic_Tsv(CurrentRexForensic_Cell(dataArr, rowIndex, mapCalc, currentColumn)) & vbTab & CurrentRexForensic_Tsv(CurrentRexForensic_Cell(dataArr, rowIndex, mapCalc, rexColumn)) & vbTab & status
    Next idKey
    Close #f

End Sub

Private Sub CurrentRexForensic_WriteSummaryRollup( _
    ByVal folderPath As String, _
    ByRef dataArr As Variant, _
    ByVal mapCalc As Object)

    Dim f As Integer
    Dim r As Long
    Dim isSummary As String
    Dim status As String

    f = FreeFile
    Open CurrentRexForensic_Path(folderPath, "summary_rollup_diff.tsv") For Output As #f
    Print #f, "Row" & vbTab & "ID" & vbTab & "WBS" & vbTab & "TF" & vbTab & "TF_REX" & vbTab & "CP" & vbTab & "CP_REX" & vbTab & "Status"
    For r = 1 To UBound(dataArr, 1)
        isSummary = UCase$(Trim$(CStr(dataArr(r, mapCalc("IsSummary")))))
        If isSummary = "YES" Or isSummary = "TRUE" Then
            If CurrentRexForensic_SameCell(dataArr, r, mapCalc, "Total Float", "Total Float REX") And CurrentRexForensic_SameCell(dataArr, r, mapCalc, "Critical Path", "Critical Path REX") Then status = "SAME" Else status = "DIFF"
            Print #f, CStr(r) & vbTab & CurrentRexForensic_Tsv(CStr(dataArr(r, mapCalc("ID")))) & vbTab & CurrentRexForensic_Tsv(CStr(dataArr(r, mapCalc("WBS")))) & vbTab & CurrentRexForensic_Tsv(CurrentRexForensic_Cell(dataArr, r, mapCalc, "Total Float")) & vbTab & CurrentRexForensic_Tsv(CurrentRexForensic_Cell(dataArr, r, mapCalc, "Total Float REX")) & vbTab & CurrentRexForensic_Tsv(CurrentRexForensic_Cell(dataArr, r, mapCalc, "Critical Path")) & vbTab & CurrentRexForensic_Tsv(CurrentRexForensic_Cell(dataArr, r, mapCalc, "Critical Path REX")) & vbTab & status
        End If
    Next r
    Close #f

End Sub

Private Sub CurrentRexForensic_WriteSynthetic(ByVal folderPath As String)

    Dim f As Integer
    f = FreeFile
    Open CurrentRexForensic_Path(folderPath, "synthetic_fixture_results.tsv") For Output As #f
    Print #f, "Fixture" & vbTab & "ExpectedUnderContract1" & vbTab & "ExpectedUnderContract2" & vbTab & "Status" & vbTab & "Notes"
    Print #f, "A_chain_no_gap" & vbTab & "Current=REX" & vbTab & "Current=REX" & vbTab & "DOCUMENTED_STATIC" & vbTab & "If baseline dates are logic-driven, both contracts converge."
    Print #f, "B_gap_without_lag" & vbTab & "Current=REX" & vbTab & "Current=REX" & vbTab & "DOCUMENTED_STATIC" & vbTab & "The reproduction workbook verifies that unmodelled baseline gaps are preserved."
    Print #f, "C_gap_with_explicit_lag" & vbTab & "Current should converge with REX if lag encodes the gap" & vbTab & "Current=REX" & vbTab & "DOCUMENTED_STATIC" & vbTab & "Use 4.3 -> 5.2.1 +30 as positive control."
    Close #f

End Sub

Private Sub CurrentRexForensic_WriteMissingColumn(ByVal folderPath As String, ByVal fileName As String, ByVal note As String)
    Dim f As Integer
    f = FreeFile
    Open CurrentRexForensic_Path(folderPath, fileName) For Output As #f
    Print #f, "Status" & vbTab & "Notes"
    Print #f, "NOT_AVAILABLE" & vbTab & CurrentRexForensic_Tsv(note)
    Close #f
End Sub

Private Function CurrentRexForensic_InvariantEligible(ByRef dataArr As Variant, ByVal mapCalc As Object, ByVal rowIndex As Long) As Boolean
    If Not CurrentRexForensic_EmptyColumn(dataArr, rowIndex, mapCalc, "Actual Start") Then Exit Function
    If Not CurrentRexForensic_EmptyColumn(dataArr, rowIndex, mapCalc, "Actual Finish") Then Exit Function
    If Not CurrentRexForensic_EmptyColumn(dataArr, rowIndex, mapCalc, "Forecast Start") Then Exit Function
    If Not CurrentRexForensic_EmptyColumn(dataArr, rowIndex, mapCalc, "Forecast Finish") Then Exit Function
    If Not CurrentRexForensic_SameCell(dataArr, rowIndex, mapCalc, "Calculated Start", "Baseline Start") Then Exit Function
    If Not CurrentRexForensic_SameCell(dataArr, rowIndex, mapCalc, "Calculated Finish", "Baseline Finish") Then Exit Function
    If Not CurrentRexForensic_SameCell(dataArr, rowIndex, mapCalc, "Calculated Duration", "Baseline Duration") Then Exit Function
    CurrentRexForensic_InvariantEligible = True
End Function

Private Function CurrentRexForensic_EmptyColumn(ByRef dataArr As Variant, ByVal rowIndex As Long, ByVal mapCalc As Object, ByVal columnName As String) As Boolean
    If Not mapCalc.Exists(columnName) Then
        CurrentRexForensic_EmptyColumn = True
    Else
        CurrentRexForensic_EmptyColumn = Not HasValue(GetCellValue(dataArr(rowIndex, mapCalc(columnName))))
    End If
End Function

Private Function CurrentRexForensic_SameCell(ByRef dataArr As Variant, ByVal rowIndex As Long, ByVal mapCalc As Object, ByVal leftCol As String, ByVal rightCol As String) As Boolean
    On Error GoTo FailSafe
    If Not mapCalc.Exists(leftCol) Or Not mapCalc.Exists(rightCol) Then Exit Function
    CurrentRexForensic_SameCell = CurrentRexForensic_SameValue(GetCellValue(dataArr(rowIndex, mapCalc(leftCol))), GetCellValue(dataArr(rowIndex, mapCalc(rightCol))))
    Exit Function
FailSafe:
    CurrentRexForensic_SameCell = False
End Function

Private Function CurrentRexForensic_SameValue(ByVal leftVal As Variant, ByVal rightVal As Variant) As Boolean
    On Error GoTo FailSafe
    If Not HasValue(leftVal) And Not HasValue(rightVal) Then
        CurrentRexForensic_SameValue = True
    ElseIf HasValue(leftVal) <> HasValue(rightVal) Then
        CurrentRexForensic_SameValue = False
    ElseIf (IsDate(leftVal) Or IsNumeric(leftVal)) And (IsDate(rightVal) Or IsNumeric(rightVal)) Then
        CurrentRexForensic_SameValue = (Abs(CurrentRexForensic_AsSerial(leftVal) - CurrentRexForensic_AsSerial(rightVal)) < 0.000001)
    Else
        CurrentRexForensic_SameValue = (Trim$(CStr(leftVal)) = Trim$(CStr(rightVal)))
    End If
    Exit Function
FailSafe:
    CurrentRexForensic_SameValue = False
End Function

Private Function CurrentRexForensic_Cell(ByRef dataArr As Variant, ByVal rowIndex As Long, ByVal mapCalc As Object, ByVal columnName As String) As String
    On Error GoTo FailSafe
    If Not mapCalc.Exists(columnName) Then
        CurrentRexForensic_Cell = ""
    Else
        CurrentRexForensic_Cell = CurrentRexForensic_ValueText(GetCellValue(dataArr(rowIndex, mapCalc(columnName))))
    End If
    Exit Function
FailSafe:
    CurrentRexForensic_Cell = "#ERR"
End Function

Private Function CurrentRexForensic_DictValue(ByVal d As Object, ByVal key As String) As String
    If d.Exists(key) Then
        CurrentRexForensic_DictValue = CurrentRexForensic_ValueText(d(key))
    Else
        CurrentRexForensic_DictValue = ""
    End If
End Function

Private Function CurrentRexForensic_AsSerial(ByVal value As Variant) As Double
    If IsDate(value) Then
        CurrentRexForensic_AsSerial = CDbl(CDate(value))
    Else
        CurrentRexForensic_AsSerial = CDbl(value)
    End If
End Function

Private Function CurrentRexForensic_ValueText(ByVal value As Variant) As String
    On Error GoTo FailSafe
    If Not HasValue(value) Then
        CurrentRexForensic_ValueText = ""
    ElseIf IsDate(value) Then
        CurrentRexForensic_ValueText = Format$(CDate(value), "yyyy-mm-dd")
    ElseIf IsNumeric(value) And CDbl(value) > 20000 And CDbl(value) < 80000 Then
        CurrentRexForensic_ValueText = Format$(CDate(CDbl(value)), "yyyy-mm-dd")
    Else
        CurrentRexForensic_ValueText = CStr(value)
    End If
    Exit Function
FailSafe:
    CurrentRexForensic_ValueText = "#ERR"
End Function

Private Function CurrentRexForensic_Tsv(ByVal value As String) As String
    value = Replace(value, vbCr, " ")
    value = Replace(value, vbLf, " ")
    value = Replace(value, vbTab, " ")
    CurrentRexForensic_Tsv = value
End Function

Private Function CurrentRexForensic_Path(ByVal folderPath As String, ByVal fileName As String) As String
    If Right$(folderPath, 1) = "\" Then
        CurrentRexForensic_Path = folderPath & fileName
    Else
        CurrentRexForensic_Path = folderPath & "\" & fileName
    End If
End Function
