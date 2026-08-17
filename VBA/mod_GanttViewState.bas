Attribute VB_Name = "mod_GanttViewState"
Option Explicit

'===============================================================================
' MODULE : mod_GanttViewState
' DOMAINE / DOMAIN : Gantt
'
' FR
' Possede l'etat runtime du domaine et ses transitions explicites.
' Ne persiste ni ne recalcule les donnees metier sauf mention contraire.
'
' EN
' Owns domain runtime state and its explicit transitions.
' Does not persist or recalculate business data unless stated otherwise.
'
' CONTRATS / CONTRACTS : SetGanttInternalWrite, GetGanttInternalWrite, SetGanttPreserveTestInputs, GetGanttPreserveTestInputs, EnsureGanttViewInitialized, GetGanttViewMode, GetGanttShowCriticalPath, GetGanttTimelineScaleMode
' CALLBACKS EXTERNES / EXTERNAL CALLBACKS : Aucun / None
'===============================================================================


' GANTT display state, toggle callbacks and view filtering.
' Shape.OnAction macro names are intentionally preserved.
'=====================================================

Private Const GANTT_SHEET As String = "GANTT"
Private Const WBS_SHEET As String = "WBS"
Private Const WBS_TABLE As String = "tbl_WBS"
Private Const FIRST_TIMELINE_COL As Long = 11
Private Const HEADER_ROW_2 As Long = 4
Private Const FIRST_TASK_ROW As Long = 5
Private Const COL_LOGIC As Long = 10
Private Const GANTT_ROW_HEIGHT_TASK As Double = 18
Private Const BTN_VIEW_BG_NAME As String = "btn_Gantt_View_BG"
Private Const BTN_VIEW_KNOB_NAME As String = "btn_Gantt_View_Knob"
Private Const BTN_SCALE_BG_NAME As String = "btn_Gantt_Scale_BG"
Private Const BTN_SCALE_KNOB_NAME As String = "btn_Gantt_Scale_Knob"
Private Const BTN_CONSTRAINT_BG_NAME As String = "btn_Gantt_Constraint_BG"
Private Const BTN_CONSTRAINT_KNOB_NAME As String = "btn_Gantt_Constraint_Knob"
Private Const BTN_CP_BG_NAME As String = "btn_Gantt_CP_BG"
Private Const BTN_CP_KNOB_NAME As String = "btn_Gantt_CP_Knob"
Private Const BTN_CP_MULTI_BG_NAME As String = "shp_GANTT_CPMode_BG"
Private Const BTN_CP_MULTI_KNOB_NAME As String = "shp_GANTT_CPMode_Knob"
Private Const GANTT_SCALE_DAY As String = "DAY"
Private Const GANTT_SCALE_WEEK As String = "WEEK"
Private Const GANTT_SCALE_MONTH As String = "MONTH"
Private Const GANTT_VIEW_DETAIL As String = "DETAIL"
Private Const GANTT_VIEW_SUMMARY As String = "SUMMARY"
Private Const GANTT_ANALYTICS_PATH_NONE As String = "NONE"
Private Const GANTT_ANALYTICS_PATH_CP As String = "CP"
Private Const GANTT_ANALYTICS_PATH_LP As String = "LP"

Private gGanttViewMode As String
Private gPreserveTestInputs As Boolean
Private gGanttInternalWrite As Boolean
Private gAnalyticsPathMode As String
Private gShowConstraints As Boolean
Private gShowConstraintsInitialized As Boolean
Private gTimelineScaleMode As String
Private gGanttUiStateBootstrapped As Boolean

'------------------------------------------------------------------------------
' FR: Met a jour un flag d'etat interne utilise par le workflow de rendu GANTT.
' EN: Updates an internal state flag used by the GANTT rendering workflow.
'------------------------------------------------------------------------------
Public Sub SetGanttInternalWrite(ByVal internalWrite As Boolean)
    gGanttInternalWrite = internalWrite
End Sub

'------------------------------------------------------------------------------
' FR: Retourne une valeur d'etat ou une reference utilisee par le rendu GANTT.
' EN: Returns a state value or reference used by GANTT rendering.
'------------------------------------------------------------------------------
Public Function GetGanttInternalWrite() As Boolean
    GetGanttInternalWrite = gGanttInternalWrite
End Function

'------------------------------------------------------------------------------
' FR: Met a jour un flag d'etat interne utilise par le workflow de rendu GANTT.
' EN: Updates an internal state flag used by the GANTT rendering workflow.
'------------------------------------------------------------------------------
Public Sub SetGanttPreserveTestInputs(ByVal preserveValue As Boolean)
    gPreserveTestInputs = preserveValue
End Sub

'------------------------------------------------------------------------------
' FR: Retourne une valeur d'etat ou une reference utilisee par le rendu GANTT.
' EN: Returns a state value or reference used by GANTT rendering.
'------------------------------------------------------------------------------
Public Function GetGanttPreserveTestInputs() As Boolean
    GetGanttPreserveTestInputs = gPreserveTestInputs
End Function

'------------------------------------------------------------------------------
' FR: Verifie et prepare une ressource GANTT requise avant le rendu ou l'interaction.
' EN: Ensures and prepares a GANTT resource required before rendering or interaction.
'------------------------------------------------------------------------------
Public Sub EnsureGanttViewInitialized()

    Dim perfScope As clsPerfScope

    Set perfScope = Profiler_BeginScope("EnsureGanttViewInitialized", "Gantt State")

    BootstrapGanttUiStateFromSheet

    If Trim$(gGanttViewMode) = "" Then
        gGanttViewMode = GANTT_VIEW_DETAIL
    End If

    If Trim$(gTimelineScaleMode) = "" Then
        gTimelineScaleMode = GANTT_SCALE_DAY
    End If

    If Trim$(gAnalyticsPathMode) = "" Then
        gAnalyticsPathMode = GANTT_ANALYTICS_PATH_NONE
    End If

    If Not gShowConstraintsInitialized Then
        gShowConstraints = True
        gShowConstraintsInitialized = True
    End If

End Sub

'------------------------------------------------------------------------------
' FR: Execute le helper Bootstrap Gantt Ui State From Sheet dans le workflow de rendu GANTT.
' EN: Runs the Bootstrap Gantt Ui State From Sheet helper in the GANTT rendering workflow.
'------------------------------------------------------------------------------
Private Sub BootstrapGanttUiStateFromSheet()

    Dim ws As Worksheet
    Dim isOn As Boolean
    Dim pathMode As String

    If gGanttUiStateBootstrapped Then Exit Sub
    gGanttUiStateBootstrapped = True

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(GANTT_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    If TryReadGanttTwoStateToggle(ws, BTN_VIEW_BG_NAME, BTN_VIEW_KNOB_NAME, isOn) Then
        If isOn Then
            gGanttViewMode = GANTT_VIEW_SUMMARY
        Else
            gGanttViewMode = GANTT_VIEW_DETAIL
        End If
    End If

    If TryReadGanttScaleToggle(ws, pathMode) Then
        gTimelineScaleMode = pathMode
    ElseIf TryReadGanttTwoStateToggle(ws, BTN_SCALE_BG_NAME, BTN_SCALE_KNOB_NAME, isOn) Then
        If isOn Then
            gTimelineScaleMode = GANTT_SCALE_WEEK
        Else
            gTimelineScaleMode = GANTT_SCALE_DAY
        End If
    End If

    If TryReadGanttTwoStateToggle(ws, BTN_CONSTRAINT_BG_NAME, BTN_CONSTRAINT_KNOB_NAME, isOn) Then
        gShowConstraints = isOn
        gShowConstraintsInitialized = True
    End If

    If TryReadGanttAnalyticsPathToggle(ws, pathMode) Then
        gAnalyticsPathMode = pathMode
    End If

End Sub

'------------------------------------------------------------------------------
' FR: Relit l'etat d'un toggle existant sur la feuille GANTT pour rehydrater l'etat UI.
' EN: Reads an existing GANTT toggle to rehydrate UI state.
'------------------------------------------------------------------------------
Private Function TryReadGanttTwoStateToggle( _
    ByVal ws As Worksheet, _
    ByVal bgName As String, _
    ByVal knobName As String, _
    ByRef isOn As Boolean) As Boolean

    Dim bg As Shape
    Dim knob As Shape

    On Error Resume Next
    Set bg = ws.Shapes(bgName)
    Set knob = ws.Shapes(knobName)
    On Error GoTo 0

    If bg Is Nothing Then Exit Function
    If knob Is Nothing Then Exit Function

    isOn = ((knob.Left + (knob.Width / 2)) >= (bg.Left + (bg.Width / 2)))
    TryReadGanttTwoStateToggle = True

End Function

'------------------------------------------------------------------------------
' FR: Relit l'etat d'un toggle existant sur la feuille GANTT pour rehydrater l'etat UI.
' EN: Reads an existing GANTT toggle to rehydrate UI state.
'------------------------------------------------------------------------------
Private Function TryReadGanttAnalyticsPathToggle( _
    ByVal ws As Worksheet, _
    ByRef pathMode As String) As Boolean

    Dim bg As Shape
    Dim knob As Shape
    Dim knobCenter As Double
    Dim ratio As Double

    On Error Resume Next
    Set bg = ws.Shapes(BTN_CP_BG_NAME)
    Set knob = ws.Shapes(BTN_CP_KNOB_NAME)
    On Error GoTo 0

    If bg Is Nothing Then Exit Function
    If knob Is Nothing Then Exit Function
    If bg.Width <= 0 Then Exit Function

    knobCenter = knob.Left + (knob.Width / 2)
    ratio = (knobCenter - bg.Left) / bg.Width

    If bg.Width < 40 Then
        If ratio >= 0.5 Then
            pathMode = GANTT_ANALYTICS_PATH_CP
        Else
            pathMode = GANTT_ANALYTICS_PATH_NONE
        End If
    ElseIf ratio >= 0.66 Then
        pathMode = GANTT_ANALYTICS_PATH_LP
    ElseIf ratio >= 0.34 Then
        pathMode = GANTT_ANALYTICS_PATH_CP
    Else
        pathMode = GANTT_ANALYTICS_PATH_NONE
    End If

    TryReadGanttAnalyticsPathToggle = True

End Function

'------------------------------------------------------------------------------
' FR: Relit l'etat d'un toggle existant sur la feuille GANTT pour rehydrater l'etat UI.
' EN: Reads an existing GANTT toggle to rehydrate UI state.
'------------------------------------------------------------------------------
Private Function TryReadGanttScaleToggle( _
    ByVal ws As Worksheet, _
    ByRef scaleMode As String) As Boolean

    Dim bg As Shape
    Dim knob As Shape
    Dim knobCenter As Double
    Dim ratio As Double

    On Error Resume Next
    Set bg = ws.Shapes(BTN_SCALE_BG_NAME)
    Set knob = ws.Shapes(BTN_SCALE_KNOB_NAME)
    On Error GoTo 0

    If bg Is Nothing Then Exit Function
    If knob Is Nothing Then Exit Function
    If bg.Width <= 0 Then Exit Function

    knobCenter = knob.Left + (knob.Width / 2)
    ratio = (knobCenter - bg.Left) / bg.Width

    If bg.Width < 40 Then
        If ratio >= 0.5 Then
            scaleMode = GANTT_SCALE_WEEK
        Else
            scaleMode = GANTT_SCALE_DAY
        End If
    ElseIf ratio >= 0.66 Then
        scaleMode = GANTT_SCALE_MONTH
    ElseIf ratio >= 0.34 Then
        scaleMode = GANTT_SCALE_WEEK
    Else
        scaleMode = GANTT_SCALE_DAY
    End If

    TryReadGanttScaleToggle = True

End Function
'------------------------------------------------------------------------------
' FR: Retourne une valeur d'etat ou une reference utilisee par le rendu GANTT.
' EN: Returns a state value or reference used by GANTT rendering.
'------------------------------------------------------------------------------
Public Function GetGanttViewMode() As String

    EnsureGanttViewInitialized
    GetGanttViewMode = gGanttViewMode

End Function

'------------------------------------------------------------------------------
' FR: Retourne une valeur d'etat ou une reference utilisee par le rendu GANTT.
' EN: Returns a state value or reference used by GANTT rendering.
'------------------------------------------------------------------------------
Public Function GetGanttShowCriticalPath() As Boolean
    EnsureGanttViewInitialized
    GetGanttShowCriticalPath = (gAnalyticsPathMode = GANTT_ANALYTICS_PATH_CP)
End Function

'------------------------------------------------------------------------------
' FR: Met a jour un flag d'etat interne utilise par le workflow de rendu GANTT.
' EN: Updates an internal state flag used by the GANTT rendering workflow.
'------------------------------------------------------------------------------
Private Sub SetGanttShowCriticalPath(ByVal showValue As Boolean)
    If showValue Then
        gAnalyticsPathMode = GANTT_ANALYTICS_PATH_CP
    Else
        gAnalyticsPathMode = GANTT_ANALYTICS_PATH_NONE
    End If
End Sub
'------------------------------------------------------------------------------
' FR: Retourne une valeur d'etat ou une reference utilisee par le rendu GANTT.
' EN: Returns a state value or reference used by GANTT rendering.
'------------------------------------------------------------------------------
Public Function GetGanttTimelineScaleMode() As String

    EnsureGanttViewInitialized
    GetGanttTimelineScaleMode = gTimelineScaleMode

End Function

'------------------------------------------------------------------------------
' FR: Met a jour un flag d'etat interne utilise par le workflow de rendu GANTT.
' EN: Updates an internal state flag used by the GANTT rendering workflow.
'------------------------------------------------------------------------------
Public Sub SetGanttTimelineScaleMode(ByVal scaleMode As String)

    Dim oldScaleMode As String
    Dim nextScaleMode As String

    EnsureGanttViewInitialized
    oldScaleMode = gTimelineScaleMode

    Select Case UCase$(Trim$(scaleMode))
        Case GANTT_SCALE_MONTH
            nextScaleMode = GANTT_SCALE_MONTH
        Case GANTT_SCALE_WEEK
            nextScaleMode = GANTT_SCALE_WEEK
        Case Else
            nextScaleMode = GANTT_SCALE_DAY
    End Select

    gTimelineScaleMode = nextScaleMode
    If oldScaleMode <> gTimelineScaleMode Then
        GanttDependency_InvalidateLocalIndex "TimelineScaleChanged"
    End If

End Sub


'------------------------------------------------------------------------------
' FR: Retourne la valeur Gantt Show Constraints sans exposer de mutateur sur l'etat source.
' EN: Returns the Gantt Show Constraints value without exposing a mutator for source state.
'------------------------------------------------------------------------------

Public Function GetGanttShowConstraints() As Boolean
    EnsureGanttViewInitialized
    GetGanttShowConstraints = gShowConstraints
End Function

'------------------------------------------------------------------------------
' FR: Retourne la valeur Gantt Analytics Path Mode sans exposer de mutateur sur l'etat source.
' EN: Returns the Gantt Analytics Path Mode value without exposing a mutator for source state.
'------------------------------------------------------------------------------

Public Function GetGanttAnalyticsPathMode() As String
    EnsureGanttViewInitialized
    GetGanttAnalyticsPathMode = gAnalyticsPathMode
End Function

'------------------------------------------------------------------------------
' FR: Bascule un mode UI GANTT puis relance le refresh adapte a ce changement.
' EN: Toggles a GANTT UI mode then runs the refresh suited to that change.
'------------------------------------------------------------------------------
Public Sub Toggle_Gantt_View()

    Dim ws As Worksheet
    Dim oldScreenUpdating As Boolean

    EnsureGanttViewInitialized
    Set ws = ThisWorkbook.Worksheets(GANTT_SHEET)

    If gGanttViewMode = GANTT_VIEW_SUMMARY Then
        gGanttViewMode = GANTT_VIEW_DETAIL
    Else
        gGanttViewMode = GANTT_VIEW_SUMMARY
    End If

    oldScreenUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False
    On Error GoTo SafeExit

    GanttDrag_PauseForLifecycle
    ApplyCurrentGanttView ws
    GanttUiControls_EnsureCanonical ws

SafeExit:
    On Error Resume Next
    GanttDrag_ReconcileWatchState
    Application.ScreenUpdating = oldScreenUpdating
    On Error GoTo 0

End Sub

'------------------------------------------------------------------------------
' FR: Bascule un mode UI GANTT puis relance le refresh adapte a ce changement.
' EN: Toggles a GANTT UI mode then runs the refresh suited to that change.
'------------------------------------------------------------------------------
Public Sub Toggle_Gantt_CriticalPath()

    EnsureGanttViewInitialized
    GanttLocal_PrimeNormalState

    Select Case gAnalyticsPathMode
        Case GANTT_ANALYTICS_PATH_NONE
            gAnalyticsPathMode = GANTT_ANALYTICS_PATH_CP
        Case GANTT_ANALYTICS_PATH_CP
            gAnalyticsPathMode = GANTT_ANALYTICS_PATH_LP
        Case Else
            gAnalyticsPathMode = GANTT_ANALYTICS_PATH_NONE
    End Select

    GanttRefresh_ApplyAnalyticsStyleOnly

End Sub

'------------------------------------------------------------------------------
' FR: Bascule un mode UI GANTT puis relance le refresh adapte a ce changement.
' EN: Toggles a GANTT UI mode then runs the refresh suited to that change.
'------------------------------------------------------------------------------
Public Sub Toggle_Gantt_Constraints()

    Dim ws As Worksheet

    EnsureGanttViewInitialized
    gShowConstraints = Not gShowConstraints
    Set ws = ThisWorkbook.Worksheets(GANTT_SHEET)

    GanttUiControls_RefreshConstraintVisual ws
    If gShowConstraints Then
        If Not GanttConstraint_RefreshCurrentOverlay(ws) Then
            Err.Raise 5, "Toggle_Gantt_Constraints", "Constraint overlay refresh failed."
        End If
    Else
        GanttConstraint_ClearCurrentOverlay ws
    End If
    GanttUiControls_RefreshConstraintVisual ws

End Sub

'------------------------------------------------------------------------------
' FR: Bascule un mode UI GANTT puis relance le refresh adapte a ce changement.
' EN: Toggles a GANTT UI mode then runs the refresh suited to that change.
'------------------------------------------------------------------------------
Public Sub Toggle_Gantt_Scale()

    Dim nextScaleMode As String

    EnsureGanttViewInitialized

    Select Case gTimelineScaleMode
        Case GANTT_SCALE_DAY
            nextScaleMode = GANTT_SCALE_WEEK
        Case GANTT_SCALE_WEEK
            nextScaleMode = GANTT_SCALE_MONTH
        Case Else
            nextScaleMode = GANTT_SCALE_DAY
    End Select

    SetGanttTimelineScaleMode nextScaleMode

    Refresh_Gantt_AfterScaleChange
    GanttDrag_ReconcileWatchState

End Sub

'------------------------------------------------------------------------------
' FR: Construit une map, un index ou une structure intermediaire consommee par le rendu GANTT.
' EN: Builds a map, index, or intermediate structure consumed by GANTT rendering.
'------------------------------------------------------------------------------
Private Function BuildGanttShapeIndex(ByVal ws As Worksheet) As Object

    Dim perfScope As clsPerfScope
    Dim shapeIndex As Object
    Dim shp As Shape

    Set perfScope = Profiler_BeginScope("BuildGanttShapeIndex", "Shape Index")
    Set shapeIndex = CreateObject("Scripting.Dictionary")

    If ws Is Nothing Then
        Set BuildGanttShapeIndex = shapeIndex
        Exit Function
    End If

    For Each shp In ws.Shapes
        If Not shapeIndex.Exists(CStr(shp.Name)) Then
            shapeIndex.Add CStr(shp.Name), shp
        End If
    Next shp

    Set BuildGanttShapeIndex = shapeIndex

End Function

'------------------------------------------------------------------------------
' FR: Execute le helper Set Shape Visibility If Exists dans le workflow de rendu GANTT.
' EN: Runs the Set Shape Visibility If Exists helper in the GANTT rendering workflow.
'------------------------------------------------------------------------------
Private Sub SetShapeVisibilityIfExists( _
    ByVal shapeIndex As Object, _
    ByVal shapeName As String, _
    ByVal isVisible As Boolean)

    Dim perfScope As clsPerfScope
    Dim shp As Shape

    Set perfScope = Profiler_BeginScope("SetShapeVisibilityIfExists", "Shape Lookup")

    If shapeIndex Is Nothing Then Exit Sub
    If Not shapeIndex.Exists(shapeName) Then Exit Sub

    Set shp = shapeIndex(shapeName)
    shp.Visible = IIf(isVisible, msoTrue, msoFalse)

End Sub

'------------------------------------------------------------------------------
' FR: Execute le helper Set Row Rendered Shapes Visibility dans le workflow de rendu GANTT.
' EN: Runs the Set Row Rendered Shapes Visibility helper in the GANTT rendering workflow.
'------------------------------------------------------------------------------
Private Sub SetRowRenderedShapesVisibility( _
    ByVal ws As Worksheet, _
    ByVal rowNum As Long, _
    ByVal isVisible As Boolean, _
    ByVal shapeIndex As Object)

    Dim perfScope As clsPerfScope

    Dim suffix As String

    Set perfScope = Profiler_BeginScope("SetRowRenderedShapesVisibility", "Shape Visibility")

    If rowNum < FIRST_TASK_ROW Then Exit Sub

    suffix = CStr(rowNum - FIRST_TASK_ROW + 1)

    SetShapeVisibilityIfExists shapeIndex, "TASK_" & suffix, isVisible
    SetShapeVisibilityIfExists shapeIndex, "TASK_" & suffix & "_P", isVisible

    SetShapeVisibilityIfExists shapeIndex, "MS_" & suffix, isVisible

    SetShapeVisibilityIfExists shapeIndex, "SUM_" & suffix & "_H", isVisible
    SetShapeVisibilityIfExists shapeIndex, "SUM_" & suffix & "_L", isVisible
    SetShapeVisibilityIfExists shapeIndex, "SUM_" & suffix & "_R", isVisible
    SetShapeVisibilityIfExists shapeIndex, "SUM_" & suffix & "_TXT", isVisible

    SetConstraintMarkerVisibilityForRow ws, suffix, isVisible, shapeIndex

End Sub

'------------------------------------------------------------------------------
' FR: Gere l'affichage ou la lecture des contraintes visuelles GANTT sans modifier le moteur de calcul.
' EN: Handles GANTT visual constraint display or lookup without changing the calculation engine.
'------------------------------------------------------------------------------
Private Sub SetConstraintMarkerVisibilityForRow( _
    ByVal ws As Worksheet, _
    ByVal suffix As String, _
    ByVal isVisible As Boolean, _
    ByVal shapeIndex As Object)

    Dim perfScope As clsPerfScope

    Dim tokens As Variant
    Dim token As Variant
    Dim i As Long

    Set perfScope = Profiler_BeginScope("SetConstraintMarkerVisibilityForRow", "Shape Visibility")

    tokens = Array("SNET", "SNLT", "FNET", "FNLT")

    For Each token In tokens
        For i = 1 To 5
            SetShapeVisibilityIfExists shapeIndex, "CSTR_" & suffix & "_" & CStr(token) & "_" & CStr(i), isVisible
        Next i
    Next token

End Sub

'------------------------------------------------------------------------------
' FR: Actualise Apply Current Gantt View sans modifier les regles metier qui produisent les donnees.
' EN: Refreshes Apply Current Gantt View without changing the business rules that produce the data.
'------------------------------------------------------------------------------

Private Sub ApplyCurrentGanttView( _
    ByVal ws As Worksheet, _
    Optional ByVal freshFullRender As Boolean = False)

    Dim perfScope As clsPerfScope

    Dim lastRow As Long
    Dim r As Long
    Dim showRow As Boolean
    Dim shp As Shape
    Dim oldScreenUpdating As Boolean
    Dim shapeIndex As Object

    Set perfScope = Profiler_BeginScope("ApplyCurrentGanttView", "Gantt UI")

    EnsureGanttViewInitialized

    oldScreenUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False
    On Error GoTo SafeExit
    lastRow = GetLastRenderedGanttRow(ws)
    If lastRow < FIRST_TASK_ROW Then GoTo SafeExit

    If GetGanttViewMode() = GANTT_VIEW_DETAIL Then

        For r = FIRST_TASK_ROW To lastRow
            If ws.rows(r).Hidden Then ws.rows(r).Hidden = False
        Next r

        If freshFullRender And Not IsAggregatedScaleMode() Then
            GanttDependencySvg_SetVisible ws, True
            ReconcileGanttViewRowsAfterFiltering ws, lastRow
            Profiler_RecordOperation "GanttFullFreshVisibilityFastPath", 1, 0#
            GoTo SafeExit
        End If

        Set shapeIndex = BuildGanttShapeIndex(ws)
        For r = FIRST_TASK_ROW To lastRow
            SetRowRenderedShapesVisibility ws, r, True, shapeIndex
        Next r

        For Each shp In ws.Shapes
            If Left$(shp.Name, 4) = "DEP_" Then
                If IsAggregatedScaleMode() Then
                    shp.Visible = msoFalse
                Else
                    shp.Visible = msoTrue
                End If
            End If
        Next shp
        GanttDependencySvg_SetVisible ws, Not IsAggregatedScaleMode()

        ReconcileGanttViewRowsAfterFiltering ws, lastRow
        GoTo SafeExit
    End If

    Set shapeIndex = BuildGanttShapeIndex(ws)
    For r = FIRST_TASK_ROW To lastRow
        showRow = ShouldShowGanttRow(ws, r)
        ws.rows(r).Hidden = Not showRow
    Next r

    For r = FIRST_TASK_ROW To lastRow
        SetRowRenderedShapesVisibility ws, r, Not ws.rows(r).Hidden, shapeIndex
    Next r

    For Each shp In ws.Shapes
        If Left$(shp.Name, 4) = "DEP_" Then
            shp.Visible = msoFalse
        End If
    Next shp
    GanttDependencySvg_SetVisible ws, False

    ReconcileGanttViewRowsAfterFiltering ws, lastRow

SafeExit:
    Application.ScreenUpdating = oldScreenUpdating

End Sub

'------------------------------------------------------------------------------
' FR: Execute le helper Get Last Rendered Gantt Row dans le workflow de rendu GANTT.
' EN: Runs the Get Last Rendered Gantt Row helper in the GANTT rendering workflow.
'------------------------------------------------------------------------------
Public Function GetLastRenderedGanttRow(ByVal ws As Worksheet) As Long

    Dim perfScope As clsPerfScope

    Dim tblWBS As ListObject

    Set perfScope = Profiler_BeginScope("GetLastRenderedGanttRow", "Excel Metadata")

    On Error GoTo Fallback

    Set tblWBS = ThisWorkbook.Worksheets(WBS_SHEET).ListObjects(WBS_TABLE)
    If Not tblWBS.DataBodyRange Is Nothing Then
        GetLastRenderedGanttRow = FIRST_TASK_ROW + tblWBS.ListRows.Count - 1
        Exit Function
    End If

Fallback:
    GetLastRenderedGanttRow = GetLastGanttRow(ws)

End Function

'------------------------------------------------------------------------------
' FR: Execute le helper Reconcile Gantt View Grid After Filtering dans le workflow de rendu GANTT.
' EN: Runs the Reconcile Gantt View Grid After Filtering helper in the GANTT rendering workflow.
'------------------------------------------------------------------------------
Private Sub ReconcileGanttViewRowsAfterFiltering(ByVal ws As Worksheet, ByVal lastRow As Long)

    Dim perfScope As clsPerfScope

    Dim r As Long

    Set perfScope = Profiler_BeginScope("ReconcileGanttViewRowsAfterFiltering", "Excel Format")

    If ws Is Nothing Then Exit Sub
    If lastRow < FIRST_TASK_ROW Then Exit Sub


    For r = FIRST_TASK_ROW To lastRow
        If Not ws.rows(r).Hidden Then
            ws.rows(r).rowHeight = GANTT_ROW_HEIGHT_TASK
        End If
    Next r

End Sub

'------------------------------------------------------------------------------
' FR: Retourne une decision de rendu ou d'etat utilisee par le workflow GANTT.
' EN: Returns a rendering or state decision used by the GANTT workflow.
'------------------------------------------------------------------------------
Private Function ShouldShowGanttRow(ByVal ws As Worksheet, ByVal rowNum As Long) As Boolean

    Dim tblWBS As ListObject
    Dim dataRow As Long
    Dim summaryDisplayVal As String


    dataRow = rowNum - FIRST_TASK_ROW + 1
    If dataRow < 1 Then Exit Function

    On Error GoTo SafeExit

    Set tblWBS = ThisWorkbook.Worksheets(WBS_SHEET).ListObjects(WBS_TABLE)
    If tblWBS.DataBodyRange Is Nothing Then Exit Function
    If dataRow > tblWBS.ListRows.Count Then Exit Function

    summaryDisplayVal = UCase$(Trim$(CStr(tblWBS.DataBodyRange.cells(dataRow, tblWBS.ListColumns("S").Index).value)))
    ShouldShowGanttRow = (summaryDisplayVal = "Y")

SafeExit:

End Function


'------------------------------------------------------------------------------
' FR: Actualise Apply Gantt Ui State sans modifier les regles metier qui produisent les donnees.
' EN: Refreshes Apply Gantt Ui State without changing the business rules that produce the data.
'------------------------------------------------------------------------------

Public Sub ApplyGanttUiState( _
    ByVal ws As Worksheet, _
    Optional ByVal rebuildHeaderControls As Boolean = True, _
    Optional ByVal freshFullRender As Boolean = False)

    Dim perfScope As clsPerfScope

    Set perfScope = Profiler_BeginScope("ApplyGanttUiState", "Gantt UI")

    'The historical flag is retained for source compatibility. Both full and fast
    'paths now share the same differential post-condition for header controls.
    GanttUiControls_EnsureCanonical ws

    If rebuildHeaderControls Then ApplyCurrentGanttView ws, freshFullRender

End Sub
'------------------------------------------------------------------------------
' FR: Met en forme ou met a jour un element UI/shape du GANTT.
' EN: Formats or updates a GANTT UI/shape element.
'------------------------------------------------------------------------------
Private Sub Refresh_Gantt_AfterScaleChange()

    Dim oldPreserve As Boolean

    On Error GoTo SafeExit

    oldPreserve = GetGanttPreserveTestInputs()
    SetGanttPreserveTestInputs True
    CommitGanttUpdate GanttEngine_ActiveDataSource(), GANTT_UPDATE_SCOPE_FULL, GANTT_RENDER_INTENT_SHOW, "Refresh_Gantt_AfterScaleChange"

SafeExit:
    SetGanttPreserveTestInputs oldPreserve

End Sub

'------------------------------------------------------------------------------
' FR: Met en forme ou met a jour un element UI/shape du GANTT.
' EN: Formats or updates a GANTT UI/shape element.
'------------------------------------------------------------------------------
Private Sub Refresh_Gantt_UI_Only()

    Dim oldPreserve As Boolean

    On Error GoTo SafeExit

    oldPreserve = GetGanttPreserveTestInputs()

    SetGanttPreserveTestInputs True
    CommitGanttUpdate GanttEngine_ActiveDataSource(), GANTT_UPDATE_SCOPE_INCREMENTAL, GANTT_RENDER_INTENT_OFFSCREEN, "Refresh_Gantt_UI_Only"

SafeExit:
    SetGanttPreserveTestInputs oldPreserve

End Sub

