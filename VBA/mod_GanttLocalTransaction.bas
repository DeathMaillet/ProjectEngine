Attribute VB_Name = "mod_GanttLocalTransaction"
Option Explicit

'===============================================================================
' MODULE : mod_GanttLocalTransaction
' DOMAIN : Gantt
'
' Owns the last committed visual model and builds workbook-local change sets.
' It never reads shape geometry and never owns planning calculations.
'===============================================================================

Private gCommittedSnapshot As Object
Private gCommittedScale As String
Private gCommittedView As String
Private gCommittedRowCount As Long
Private gSnapshotVersion As Long

Public Sub GanttLocal_Invalidate(Optional ByVal reason As String = "")

    Set gCommittedSnapshot = Nothing
    gCommittedScale = ""
    gCommittedView = ""
    gCommittedRowCount = 0
    gSnapshotVersion = gSnapshotVersion + 1

    If Len(Trim$(reason)) > 0 Then
        Profiler_RecordOperation "GanttLocalInvalidation_" & reason, 1, 0#
    End If

End Sub

Public Sub GanttLocal_ForgetIds( _
    ByVal ids As Object, _
    Optional ByVal reason As String = "")

    Dim key As Variant
    Dim removedCount As Long

    If ids Is Nothing Then Exit Sub
    If gCommittedSnapshot Is Nothing Then Exit Sub

    For Each key In ids.keys
        If gCommittedSnapshot.Exists(CStr(key)) Then
            gCommittedSnapshot.Remove CStr(key)
            removedCount = removedCount + 1
        End If
    Next key

    If removedCount > 0 Then
        gSnapshotVersion = gSnapshotVersion + 1
        Profiler_RecordOperation "GanttLocalForgotIds", removedCount, 0#
        If Len(Trim$(reason)) > 0 Then
            Profiler_RecordOperation "GanttLocalForgotIds_" & reason, removedCount, 0#
        End If
    End If

End Sub

Public Function GanttLocal_HasCommittedSnapshot() As Boolean

    GanttLocal_HasCommittedSnapshot = Not gCommittedSnapshot Is Nothing

End Function

Public Function GanttLocal_PrimeNormalState() As Boolean

    Dim perfScope As clsPerfScope
    Dim wsWBS As Worksheet
    Dim wsGantt As Worksheet
    Dim tblWBS As ListObject
    Dim dataArr As Variant
    Dim mapWBS As Object
    Dim hasChildren As Object
    Dim baseById As Object
    Dim emptySimulation As Object
    Dim changeSet As Object

    Set perfScope = Profiler_BeginScope("GanttLocal_PrimeNormalState", "Gantt Local")
    On Error GoTo Failed

    If Not gCommittedSnapshot Is Nothing Then
        GanttLocal_PrimeNormalState = True
        Exit Function
    End If

    Set wsWBS = ThisWorkbook.Worksheets("WBS")
    Set wsGantt = ThisWorkbook.Worksheets("GANTT")
    Set tblWBS = wsWBS.ListObjects("tbl_WBS")
    If tblWBS.DataBodyRange Is Nothing Then Exit Function

    Set mapWBS = CanonicalIdentity_BuildColumnMap(tblWBS)
    ValidateGanttSourceColumns mapWBS
    dataArr = tblWBS.DataBodyRange.Value
    Set hasChildren = GanttHierarchy_BuildDirectParentPresenceFromWbs(dataArr, mapWBS)
    Set baseById = GanttLive_BuildBaseByIdMap()
    Set emptySimulation = CreateObject("Scripting.Dictionary")

    Set changeSet = GanttLocal_BuildChangeSet( _
        dataArr, mapWBS, hasChildren, baseById, emptySimulation, False, "PRIME")
    GanttLocal_CommitChangeSet changeSet

    If Not GanttDependency_PrimeLocalIndex(wsGantt) Then
        GanttLocal_Invalidate "DependencyPrimeFailed"
        Exit Function
    End If

    Profiler_RecordOperation "GanttLocalPrimeSuccess", 1, 0#
    GanttLocal_PrimeNormalState = True
    Exit Function

Failed:
    GanttLocal_Invalidate "PrimeError"
    Profiler_RecordOperation "GanttLocalPrimeError", 1, 0#

End Function

Public Function GanttLocal_BuildChangeSet( _
    ByRef dataArr As Variant, _
    ByVal mapWBS As Object, _
    ByVal hasChildren As Object, _
    ByVal baseById As Object, _
    ByVal testById As Object, _
    ByVal isTestMode As Boolean, _
    ByVal transactionType As String) As Object

    Dim perfScope As clsPerfScope
    Dim changeSet As Object
    Dim changedIds As Object
    Dim changedRows As Object
    Dim currentSnapshot As Object
    Dim wbsToId As Object
    Dim idToWbs As Object
    Dim idToRow As Object
    Dim idVal As String
    Dim wbs As String
    Dim parentWbs As String
    Dim parentId As String
    Dim rowSignature As String
    Dim r As Long
    Dim rowCount As Long
    Dim key As Variant
    Dim changedKeys As Variant
    Dim parentRow As Long
    Dim fallbackReason As String
    Dim localEligible As Boolean
    Dim topologyChanged As Boolean

    Set perfScope = Profiler_BeginScope("GanttLocal_BuildChangeSet", "Gantt Local")
    Set changeSet = CreateObject("Scripting.Dictionary")
    Set changedIds = CreateObject("Scripting.Dictionary")
    Set changedRows = CreateObject("Scripting.Dictionary")
    Set currentSnapshot = CreateObject("Scripting.Dictionary")
    Set wbsToId = CreateObject("Scripting.Dictionary")
    Set idToWbs = CreateObject("Scripting.Dictionary")
    Set idToRow = CreateObject("Scripting.Dictionary")

    rowCount = UBound(dataArr, 1)

    For r = 1 To rowCount
        idVal = Trim$(CStr(dataArr(r, mapWBS(VTS_COL_ID))))
        wbs = NormalizeWBS(CStr(dataArr(r, mapWBS(VTS_COL_WBS))))

        If idVal <> "" Then
            rowSignature = GanttLocal_BuildRowSignature( _
                dataArr, mapWBS, hasChildren, r, idVal, wbs, _
                baseById, testById, isTestMode)

            currentSnapshot(idVal) = rowSignature
            idToWbs(idVal) = wbs
            idToRow(idVal) = r
            If wbs <> "" Then wbsToId(wbs) = idVal

            If gCommittedSnapshot Is Nothing Then
                If isTestMode Then
                    If GanttLive_HasRenderableTestDelta(idVal, baseById, testById) Then
                        changedIds(idVal) = True
                        changedRows(CStr(r)) = True
                    End If
                Else
                    changedIds(idVal) = True
                    changedRows(CStr(r)) = True
                End If
            ElseIf Not gCommittedSnapshot.Exists(idVal) Then
                changedIds(idVal) = True
                changedRows(CStr(r)) = True
            ElseIf CStr(gCommittedSnapshot(idVal)) <> rowSignature Then
                changedIds(idVal) = True
                changedRows(CStr(r)) = True
                If GanttLocal_IsPredecessorSignatureChange( _
                    CStr(gCommittedSnapshot(idVal)), rowSignature) Then
                    topologyChanged = True
                End If
                If GanttLocal_IsStructuralSignatureChange( _
                    CStr(gCommittedSnapshot(idVal)), rowSignature) Then
                    fallbackReason = "StructuralRowChanged"
                End If
            End If
        End If
    Next r

    If Not gCommittedSnapshot Is Nothing Then
        For Each key In gCommittedSnapshot.Keys
            If Not currentSnapshot.Exists(CStr(key)) Then
                fallbackReason = "DatasetChanged"
                Exit For
            End If
        Next key
    End If

    'A child progress/date change can alter a summary even when the summary row
    'has not yet changed textually. Include all visible WBS ancestors.
    If changedIds.Count > 0 Then
        changedKeys = changedIds.Keys
        For Each key In changedKeys
            If idToWbs.Exists(CStr(key)) Then
                wbs = CStr(idToWbs(CStr(key)))
                parentWbs = GanttLocal_ParentWbs(wbs)

                Do While parentWbs <> ""
                    If wbsToId.Exists(parentWbs) Then
                        parentId = CStr(wbsToId(parentWbs))
                        changedIds(parentId) = True
                        parentRow = 0
                        If idToRow.Exists(parentId) Then parentRow = CLng(idToRow(parentId))
                        If parentRow > 0 Then changedRows(CStr(parentRow)) = True
                    End If
                    parentWbs = GanttLocal_ParentWbs(parentWbs)
                Loop
            End If
        Next key
    End If

    If gCommittedSnapshot Is Nothing Then
        If isTestMode Then
            Profiler_RecordOperation "GanttLocalFirstSimulationCommitWithoutSnapshot", changedIds.Count, 0#
        Else
            fallbackReason = "RegistryAbsent"
        End If
    End If
    If fallbackReason = "" And gCommittedRowCount <> rowCount Then fallbackReason = "DatasetChanged"
    If fallbackReason = "" And gCommittedScale <> GetGanttTimelineScaleMode() Then fallbackReason = "ScaleChanged"
    If fallbackReason = "" And gCommittedView <> GetGanttViewMode() Then fallbackReason = "ViewChanged"

    localEligible = (fallbackReason = "")
    If localEligible Then
        Profiler_RecordOperation "GanttLocalCostModelAccepted", changedIds.Count, 0#
    End If

    changeSet.Add "TransactionType", UCase$(Trim$(transactionType))
    changeSet.Add "ChangedIds", changedIds
    changeSet.Add "ChangedRows", changedRows
    changeSet.Add "NewSnapshot", currentSnapshot
    changeSet.Add "RecordCount", rowCount
    changeSet.Add "ChangedCount", changedIds.Count
    changeSet.Add "LocalEligible", localEligible
    changeSet.Add "FallbackReason", fallbackReason
    changeSet.Add "TopologyChanged", topologyChanged
    changeSet.Add "SnapshotVersion", gSnapshotVersion + 1

    Profiler_RecordOperation "GanttLocalRecordsCompared", rowCount, 0#
    Profiler_RecordOperation "GanttLocalRecordsChanged", changedIds.Count, 0#
    If localEligible Then
        Profiler_RecordOperation "GanttLocalTransactions", 1, 0#
    Else
        Profiler_RecordOperation "GanttLocalFallback_" & fallbackReason, 1, 0#
    End If

    Set GanttLocal_BuildChangeSet = changeSet

End Function

Public Sub GanttLocal_CommitChangeSet(ByVal changeSet As Object)

    If changeSet Is Nothing Then Exit Sub
    If Not changeSet.Exists("NewSnapshot") Then Exit Sub

    Set gCommittedSnapshot = changeSet("NewSnapshot")
    gCommittedScale = GetGanttTimelineScaleMode()
    gCommittedView = GetGanttViewMode()
    gCommittedRowCount = CLng(changeSet("RecordCount"))
    gSnapshotVersion = CLng(changeSet("SnapshotVersion"))

    Profiler_RecordOperation "GanttLocalSnapshotCommits", 1, 0#

End Sub

Public Sub GanttLocal_CaptureFullSnapshot( _
    ByRef dataArr As Variant, _
    ByVal mapWBS As Object, _
    ByVal hasChildren As Object, _
    ByVal baseById As Object, _
    ByVal testById As Object, _
    ByVal isTestMode As Boolean, _
    Optional ByVal transactionType As String = "FULL")

    Dim changeSet As Object

    Set changeSet = GanttLocal_BuildChangeSet( _
        dataArr, mapWBS, hasChildren, baseById, testById, isTestMode, transactionType)
    GanttLocal_CommitChangeSet changeSet

End Sub

Private Function GanttLocal_BuildRowSignature( _
    ByRef dataArr As Variant, _
    ByVal mapWBS As Object, _
    ByVal hasChildren As Object, _
    ByVal rowIndex As Long, _
    ByVal idVal As String, _
    ByVal wbs As String, _
    ByVal baseById As Object, _
    ByVal testById As Object, _
    ByVal isTestMode As Boolean) As String

    Dim startVal As Variant
    Dim finishVal As Variant
    Dim durationVal As Variant
    Dim progressVal As Variant
    Dim isParent As Boolean
    Dim isMilestone As Boolean
    Dim isLoE As Boolean
    Dim isHighlighted As Boolean
    Dim separator As String

    startVal = GanttLive_GetDisplayStart(idVal, baseById, testById, isTestMode)
    finishVal = GanttLive_GetDisplayFinish(idVal, baseById, testById, isTestMode)
    durationVal = GanttLive_GetDisplayDuration(idVal, baseById, testById, isTestMode)
    progressVal = GanttLive_GetDisplayProgress(idVal, baseById, testById, isTestMode)

    isParent = hasChildren.Exists(wbs)
    isMilestone = TaskTypeRules_IsMilestoneRow(dataArr, mapWBS, rowIndex, VTS_COL_TASK_TYPE)
    isLoE = TaskTypeRules_IsLevelOfEffortRow(dataArr, mapWBS, rowIndex, VTS_COL_TASK_TYPE)
    isHighlighted = ShouldHighlightGanttAnalyticsPath( _
        dataArr, mapWBS, rowIndex, idVal, testById, isTestMode)
    separator = Chr$(30)

    GanttLocal_BuildRowSignature = _
        CStr(rowIndex) & separator & idVal & separator & wbs & separator & _
        GanttLocal_ValueSignature(startVal) & separator & _
        GanttLocal_ValueSignature(finishVal) & separator & _
        GanttLocal_ValueSignature(durationVal) & separator & _
        GanttLocal_ValueSignature(progressVal) & separator & _
        CStr(isParent) & separator & CStr(isMilestone) & separator & _
        CStr(isLoE) & separator & CStr(isHighlighted) & separator & _
        Trim$(CStr(dataArr(rowIndex, mapWBS(VTS_COL_PREDECESSORS_WBS)))) & separator & _
        Trim$(CStr(dataArr(rowIndex, mapWBS(VTS_COL_TASK_NAME))))

End Function

Private Function GanttLocal_IsStructuralSignatureChange( _
    ByVal oldSignature As String, _
    ByVal newSignature As String) As Boolean

    Dim oldParts As Variant
    Dim newParts As Variant

    oldParts = Split(oldSignature, Chr$(30))
    newParts = Split(newSignature, Chr$(30))

    If UBound(oldParts) < 11 Or UBound(newParts) < 11 Then
        GanttLocal_IsStructuralSignatureChange = True
        Exit Function
    End If

    GanttLocal_IsStructuralSignatureChange = _
        (CStr(oldParts(0)) <> CStr(newParts(0))) Or _
        (CStr(oldParts(1)) <> CStr(newParts(1))) Or _
        (CStr(oldParts(2)) <> CStr(newParts(2))) Or _
        (CStr(oldParts(7)) <> CStr(newParts(7))) Or _
        (CStr(oldParts(8)) <> CStr(newParts(8))) Or _
        (CStr(oldParts(9)) <> CStr(newParts(9)))

End Function

Private Function GanttLocal_IsPredecessorSignatureChange( _
    ByVal oldSignature As String, _
    ByVal newSignature As String) As Boolean

    Dim oldParts As Variant
    Dim newParts As Variant

    oldParts = Split(oldSignature, Chr$(30))
    newParts = Split(newSignature, Chr$(30))
    If UBound(oldParts) < 11 Or UBound(newParts) < 11 Then Exit Function

    GanttLocal_IsPredecessorSignatureChange = _
        (CStr(oldParts(11)) <> CStr(newParts(11)))

End Function

Private Function GanttLocal_ValueSignature(ByVal value As Variant) As String

    If IsError(value) Then
        GanttLocal_ValueSignature = "#ERR"
    ElseIf IsEmpty(value) Or IsNull(value) Then
        GanttLocal_ValueSignature = "#EMPTY"
    ElseIf IsDate(value) Then
        GanttLocal_ValueSignature = Format$(CDbl(CDate(value)), "0.000000")
    ElseIf IsNumeric(value) Then
        GanttLocal_ValueSignature = Format$(CDbl(value), "0.000000")
    Else
        GanttLocal_ValueSignature = CStr(value)
    End If

End Function

Private Function GanttLocal_ParentWbs(ByVal wbs As String) As String

    Dim dotPos As Long

    dotPos = InStrRev(wbs, ".")
    If dotPos > 0 Then GanttLocal_ParentWbs = Left$(wbs, dotPos - 1)

End Function
