Attribute VB_Name = "mod_Dashboard"
Option Explicit

'===============================================================================
' MODULE : mod_Dashboard
' DOMAINE / DOMAIN : Dashboard Rendering
'
' FR
' Construit et actualise les KPI, textes, graphiques et controles du Dashboard.
' Ne possede pas les calculs Core ou S-Curve sources.
'
' EN
' Builds and refreshes Dashboard KPIs, text, charts and controls.
' Does not own source Core or S-Curve calculations.
'
' CONTRATS / CONTRACTS : Refresh_Dashboard, Refresh_Dashboard_FullBuild, Refresh_Dashboard_ContentOnly, Dashboard_SetLanguage, Dashboard_ApplyLanguage, Toggle_Dashboard_Language, Reset_Dashboard, Run_Dashboard_Update
' CALLBACKS EXTERNES / EXTERNAL CALLBACKS : Aucun / None
'===============================================================================


Private Const DASHBOARD_SHEET As String = "DASHBOARD"
Private Const DASHBOARD_SNAPSHOT_SHEET As String = "DASHBOARD_SNAPSHOTS"
Private Const DASHBOARD_SNAPSHOT_TABLE As String = "tbl_DASHBOARD_SNAPSHOTS"
Private Const DASH_CHART_SCURVE As String = "cht_Dashboard_SCurve"
Private Const DASH_SCURVE_TODAY_LINE As String = "DASHBOARD_SCURVE_TODAY_LINE"
Private Const DASH_PREFIX As String = "DASH_"
Private Const DASH_FROM_CELL As String = "C4"
Private Const DASH_TO_CELL As String = "G4"
Private gDashboardLanguage As String
Private gHotSpotHelper As String
Private gHotSpotStep As String

'------------------------------------------------------------------------------
' FR: Rafraichit Dashboard a partir de l'etat courant.
' EN: Refreshes Dashboard from the current state.
'------------------------------------------------------------------------------
Public Sub Refresh_Dashboard(Optional ByVal preserveSnapshotSelection As Boolean = False)

    Dashboard_SetLanguage Settings_GetOwnerLanguage("DASHBOARD")
    If Dashboard_HasUsableLayout() Then
        Refresh_Dashboard_ContentOnly preserveSnapshotSelection
    Else
        Refresh_Dashboard_FullBuild preserveSnapshotSelection
    End If

End Sub
'------------------------------------------------------------------------------
' FR: Rafraichit Dashboard Full Build a partir de l'etat courant.
' EN: Refreshes Dashboard Full Build from the current state.
'------------------------------------------------------------------------------
Public Sub Refresh_Dashboard_FullBuild(Optional ByVal preserveSnapshotSelection As Boolean = False)

    Dim ws As Worksheet
    Dim readContext As clsDashboardReadContext
    Dim fromLabel As String
    Dim toLabel As String
    Dim langCode As String

    On Error GoTo ErrHandler
    Dashboard_SetLanguage Settings_GetOwnerLanguage("DASHBOARD")
    Set ws = Dashboard_EnsureSheet()
    If preserveSnapshotSelection Then
        fromLabel = CStr(ws.Range(DASH_FROM_CELL).value)
        toLabel = CStr(ws.Range(DASH_TO_CELL).value)
    End If
    langCode = Dashboard_CurrentLanguage()
    Dashboard_Clear ws
    Dashboard_SetupCanvas ws

    Set readContext = DashboardReadContext_Load()

    Dashboard_RenderHeader ws, fromLabel, toLabel
    Dashboard_RenderExecutiveSummary ws, readContext.WbsTable, readContext.WbsColumns, readContext.CalcTable, readContext.CalcColumns, readContext.SCurveProjection
    Dashboard_RenderSCurveChart ws, readContext.SCurveProjection
    Dashboard_RenderPlanningOverview ws, readContext.WbsTable, readContext.WbsColumns, readContext.CalcTable, readContext.CalcColumns
    Dashboard_RenderHotSpots ws, readContext.WbsTable, readContext.WbsColumns, readContext.CalcTable, readContext.CalcColumns

SafeExit:
    Exit Sub
ErrHandler:
    Debug.Print "Refresh_Dashboard_FullBuild error " & Err.Number & ": " & Err.Description
    Resume SafeExit
End Sub

'------------------------------------------------------------------------------
' FR: Rafraichit Dashboard Content Only a partir de l'etat courant.
' EN: Refreshes Dashboard Content Only from the current state.
'------------------------------------------------------------------------------
Public Sub Refresh_Dashboard_ContentOnly(Optional ByVal preserveSnapshotSelection As Boolean = False)

    Dim ws As Worksheet
    Dim readContext As clsDashboardReadContext
    Dim fromLabel As String
    Dim toLabel As String
    Dim oldScreenUpdating As Boolean
    On Error GoTo ErrHandler
    Dashboard_SetLanguage Settings_GetOwnerLanguage("DASHBOARD")

    If Not Dashboard_HasUsableLayout() Then
        Refresh_Dashboard_FullBuild preserveSnapshotSelection
        Exit Sub
    End If

    Set ws = ThisWorkbook.Worksheets(DASHBOARD_SHEET)
    If preserveSnapshotSelection Then
        fromLabel = CStr(ws.Range(DASH_FROM_CELL).value)
        toLabel = CStr(ws.Range(DASH_TO_CELL).value)
    End If
    oldScreenUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False

    Set readContext = DashboardReadContext_Load()
    Dashboard_UpdateHeaderTexts ws
    Dashboard_UpdateHeaderTimestamp ws
    Dashboard_SetupSnapshotControls ws, fromLabel, toLabel, False
    Dashboard_EnsureResetButton ws
    Dashboard_RemoveLegacyLanguageToggle ws
    Dashboard_UpdateKnownShapeTexts ws
    Dashboard_UpdateKpiCardsInPlace ws, readContext.WbsTable, readContext.WbsColumns, readContext.CalcTable, readContext.CalcColumns, readContext.SCurveProjection
    Dashboard_UpdateSCurveChartInPlace ws, readContext.SCurveProjection
    Dashboard_UpdatePlanningOverviewContent ws, readContext.WbsTable, readContext.WbsColumns, readContext.CalcTable, readContext.CalcColumns
    Dashboard_UpdateHotSpotsContent ws, readContext.WbsTable, readContext.WbsColumns, readContext.CalcTable, readContext.CalcColumns

SafeExit:
    Application.ScreenUpdating = oldScreenUpdating
    Exit Sub
ErrHandler:
    Debug.Print "Refresh_Dashboard_ContentOnly error " & Err.Number & ": " & Err.Description
    Resume SafeExit
End Sub
'------------------------------------------------------------------------------
' FR: Active ou initialise Set Language dans l'etat runtime du composant.
' EN: Activates or initializes Set Language in the component runtime state.
'------------------------------------------------------------------------------

Public Sub Dashboard_SetLanguage(ByVal languageCode As String)

    Select Case UCase$(Trim$(languageCode))
        Case "FR"
            gDashboardLanguage = "FR"
        Case "EN"
            gDashboardLanguage = "EN"
        Case Else
            gDashboardLanguage = "EN"
    End Select

End Sub

'------------------------------------------------------------------------------
' FR: Actualise Apply Language sans modifier les regles metier qui produisent les donnees.
' EN: Refreshes Apply Language without changing the business rules that produce the data.
'------------------------------------------------------------------------------

Public Sub Dashboard_ApplyLanguage(Optional ByVal languageCode As String = "")

    If Trim$(languageCode) <> "" Then Dashboard_SetLanguage languageCode
    Refresh_Dashboard_TextsAndComparisonOnly

End Sub

'------------------------------------------------------------------------------
' FR: Bascule l'etat Dashboard Language et met a jour les sorties associees.
' EN: Toggles Dashboard Language state and updates related outputs.
'------------------------------------------------------------------------------
Public Sub Toggle_Dashboard_Language()
    Settings_ToggleDashboardLanguage
End Sub

'------------------------------------------------------------------------------
' FR: Reinitialise Reset Dashboard dans le perimetre possede par le composant.
' EN: Resets Reset Dashboard within the state owned by the component.
'------------------------------------------------------------------------------

Public Sub Reset_Dashboard()

    Dim ws As Worksheet
    Dim oldScreenUpdating As Boolean
    Dim oldEvents As Boolean

    On Error GoTo SafeExit
    Dashboard_SetLanguage Settings_GetOwnerLanguage("DASHBOARD")

    oldScreenUpdating = Application.ScreenUpdating
    oldEvents = Application.EnableEvents
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Set ws = Dashboard_EnsureSheet()

    Dashboard_ClearSnapshotsForReset
    Dashboard_ResetVisualCaches
    Dashboard_Clear ws
    Dashboard_SetupCanvas ws
    Dashboard_RenderEmptyShell ws

SafeExit:
    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreenUpdating

End Sub
'------------------------------------------------------------------------------
' FR: Lance le workflow Dashboard Update.
' EN: Runs the Dashboard Update workflow.
'------------------------------------------------------------------------------
Public Sub Run_Dashboard_Update()

    Dim workflowStarted As Boolean
    Dim stopOnlyStarted As Boolean
    Dim consoleMessages As Collection
On Error GoTo ErrHandler

    workflowStarted = EnsurePlanningWorkflowStarted("Run_Dashboard_Update")
    BeginPlanningEventRun "Run_Dashboard_Update"
    BeginPlanningWorkflowStopOnlyDisplay
    stopOnlyStarted = True
    Run_Full_Update

    If CalcEngine_HasBlockingErrorsForState() Then GoTo CleanExit
    If IsMacroAbortRequested() Then GoTo CleanExit
    Dashboard_CreateSnapshotFromCurrentData
    Refresh_Dashboard_ContentOnly
    ThisWorkbook.Worksheets(DASHBOARD_SHEET).Activate

CleanExit:
    If stopOnlyStarted Then
        EndPlanningWorkflowStopOnlyDisplay
        stopOnlyStarted = False
    End If
    If IsMacroAbortRequested() Then ShowAbortMessageOnce
    If workflowStarted Then EndPlanningWorkflow
    Exit Sub

ErrHandler:
    If stopOnlyStarted Then
        EndPlanningWorkflowStopOnlyDisplay
        stopOnlyStarted = False
    End If

    Set consoleMessages = New Collection
    CalcBridge_AddConsoleMessage consoleMessages, "STOP", _
        BiMsg( _
            "Erreur dans Run_Dashboard_Update" & vbCrLf & _
            "-> " & Err.Description, _
            "Error in Run_Dashboard_Update" & vbCrLf & _
            "-> " & Err.Description)
    CalcBridge_ShowPlanningConsole consoleMessages
    Resume CleanExit

End Sub

'------------------------------------------------------------------------------
' FR: Cree ou remet en conformite la reference Sheet de maniere idempotente.
' EN: Creates or restores the Sheet reference to a compliant state idempotently.
'------------------------------------------------------------------------------

Private Function Dashboard_EnsureSheet() As Worksheet

    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(DASHBOARD_SHEET)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = DASHBOARD_SHEET
    End If

    Set Dashboard_EnsureSheet = ws

End Function

'------------------------------------------------------------------------------
' FR: Indique si la reference Usable Layout satisfait la condition attendue, sans modifier les donnees source.
' EN: Returns whether the Usable Layout reference satisfies the expected condition without mutating source data.
'------------------------------------------------------------------------------

Private Function Dashboard_HasUsableLayout() As Boolean

    Dim ws As Worksheet
    Dim chartObj As ChartObject

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(DASHBOARD_SHEET)
    If ws Is Nothing Then Exit Function

    Set chartObj = ws.ChartObjects(DASH_CHART_SCURVE)
    If chartObj Is Nothing Then Exit Function

    If Dashboard_GetShape(ws, "DASH_KPI_1_BG") Is Nothing Then Exit Function
    If Dashboard_GetShape(ws, "DASH_KPI_1_ACCENT") Is Nothing Then Exit Function
    If Dashboard_GetShape(ws, "DASH_KPI_1_TITLE") Is Nothing Then Exit Function
    If Dashboard_GetShape(ws, "DASH_KPI_1_HERO") Is Nothing Then Exit Function
    If Dashboard_GetShape(ws, "DASH_KPI_1_SUB") Is Nothing Then Exit Function

    Dashboard_HasUsableLayout = True
    On Error GoTo 0

End Function

'------------------------------------------------------------------------------
' FR: Reinitialise Clear dans le perimetre possede par le composant.
' EN: Resets Clear within the state owned by the component.
' FR - Effet de bord : efface uniquement les donnees ou objets cibles du contrat.
' EN - Side effect: clears only data or objects targeted by the contract.
'------------------------------------------------------------------------------

Private Sub Dashboard_Clear(ByVal ws As Worksheet)

    Dim i As Long

    ws.Cells.Clear

    For i = ws.ChartObjects.Count To 1 Step -1
        ws.ChartObjects(i).Delete
    Next i

    For i = ws.Shapes.Count To 1 Step -1
        ws.Shapes(i).Delete
    Next i

End Sub

'------------------------------------------------------------------------------
' FR: Active ou initialise Setup Canvas dans l'etat runtime du composant.
' EN: Activates or initializes Setup Canvas in the component runtime state.
'------------------------------------------------------------------------------

Private Sub Dashboard_SetupCanvas(ByVal ws As Worksheet)

    Dim i As Long

    ws.Cells.Interior.Color = RGB(246, 248, 251)
    ws.Cells.Font.Name = "Segoe UI"
    ws.Cells.Font.Size = 10

    For i = 1 To 24
        ws.Columns(i).ColumnWidth = 10
    Next i

    ws.Columns("A").ColumnWidth = 2
    ws.Columns("B:C").ColumnWidth = 11
    ws.Columns("D:F").ColumnWidth = 12
    ws.Columns("G:H").ColumnWidth = 11
    ws.Columns("I:N").ColumnWidth = 12
    ws.Columns("O:W").ColumnWidth = 11

    For i = 1 To 62
        ws.rows(i).rowHeight = 18
    Next i

    ws.Range("A1:W62").Interior.Color = RGB(246, 248, 251)
    ActiveWindow.DisplayGridlines = False

    With ws.PageSetup
        .PrintArea = ""
    End With

End Sub

'------------------------------------------------------------------------------
' FR: Rafraichit Dashboard Comparison a partir de l'etat courant.
' EN: Refreshes Dashboard Comparison from the current state.
'------------------------------------------------------------------------------
Public Sub Refresh_Dashboard_Comparison()

    On Error GoTo ErrHandler

    Refresh_Dashboard_TextsAndComparisonOnly
    ThisWorkbook.Worksheets(DASHBOARD_SHEET).Activate
    Exit Sub

ErrHandler:
    MsgBox Dashboard_L("Erreur dans Refresh_Dashboard_Comparison", "Error in Refresh_Dashboard_Comparison") & vbCrLf & Err.Description, vbExclamation

End Sub

'------------------------------------------------------------------------------
' FR: Rafraichit Dashboard Texts And Comparison Only a partir de l'etat courant.
' EN: Refreshes Dashboard Texts And Comparison Only from the current state.
'------------------------------------------------------------------------------
Private Sub Refresh_Dashboard_TextsAndComparisonOnly()

    Dim ws As Worksheet
    Dim readContext As clsDashboardReadContext
    Dim fromLabel As String
    Dim toLabel As String

    On Error GoTo ErrHandler
    Set ws = ThisWorkbook.Worksheets(DASHBOARD_SHEET)
    fromLabel = CStr(ws.Range(DASH_FROM_CELL).value)
    toLabel = CStr(ws.Range(DASH_TO_CELL).value)
    Set readContext = DashboardReadContext_Load()

    Dashboard_UpdateHeaderTexts ws
    Dashboard_UpdateHeaderTimestamp ws
    Dashboard_SetupSnapshotControls ws, fromLabel, toLabel, False
    Dashboard_EnsureResetButton ws
    Dashboard_RemoveLegacyLanguageToggle ws
    Dashboard_UpdateKnownShapeTexts ws
    Dashboard_UpdateKpiCardsInPlace ws, readContext.WbsTable, readContext.WbsColumns, readContext.CalcTable, readContext.CalcColumns, readContext.SCurveProjection
    Dashboard_UpdateSCurveChartInPlace ws, readContext.SCurveProjection
    Dashboard_UpdatePlanningOverviewContent ws, readContext.WbsTable, readContext.WbsColumns, readContext.CalcTable, readContext.CalcColumns
    Dashboard_UpdateHotSpotsContent ws, readContext.WbsTable, readContext.WbsColumns, readContext.CalcTable, readContext.CalcColumns

SafeExit:
    Exit Sub
ErrHandler:
    Debug.Print "Refresh_Dashboard_TextsAndComparisonOnly error " & Err.Number & ": " & Err.Description
    Err.Raise Err.Number, "Refresh_Dashboard_TextsAndComparisonOnly", Err.Description
End Sub

'------------------------------------------------------------------------------
' FR: Reinitialise Clear Snapshots For Reset dans le perimetre possede par le composant.
' EN: Resets Clear Snapshots For Reset within the state owned by the component.
' FR - Effet de bord : efface uniquement les donnees ou objets cibles du contrat.
' EN - Side effect: clears only data or objects targeted by the contract.
'------------------------------------------------------------------------------

Private Sub Dashboard_ClearSnapshotsForReset()

    Dim tbl As ListObject

    Set tbl = Dashboard_EnsureSnapshotsTable()

    If Not tbl Is Nothing Then
        Do While tbl.ListRows.Count > 0
            tbl.ListRows(tbl.ListRows.Count).Delete
        Loop
    End If

    On Error Resume Next
    ThisWorkbook.names("DashboardSnapshotLabels").Delete
    On Error GoTo 0

End Sub

'------------------------------------------------------------------------------
' FR: Reinitialise Reset Visual Caches dans le perimetre possede par le composant.
' EN: Resets Reset Visual Caches within the state owned by the component.
'------------------------------------------------------------------------------

Private Sub Dashboard_ResetVisualCaches()

    gHotSpotHelper = vbNullString
    gHotSpotStep = vbNullString

End Sub

'------------------------------------------------------------------------------
' FR: Actualise Render Empty Shell sans modifier les regles metier qui produisent les donnees.
' EN: Refreshes Render Empty Shell without changing the business rules that produce the data.
'------------------------------------------------------------------------------

Private Sub Dashboard_RenderEmptyShell(ByVal ws As Worksheet)

    Dashboard_RenderHeader ws
    Dashboard_RenderExecutiveSummaryEmpty ws
    Dashboard_RenderSCurveEmpty ws
    Dashboard_RenderPlanningOverviewEmpty ws
    Dashboard_RenderHotSpotsEmpty ws

End Sub

'------------------------------------------------------------------------------
' FR: Actualise Render Executive Summary Empty sans modifier les regles metier qui produisent les donnees.
' EN: Refreshes Render Executive Summary Empty without changing the business rules that produce the data.
'------------------------------------------------------------------------------

Private Sub Dashboard_RenderExecutiveSummaryEmpty(ByVal ws As Worksheet)

    Dim leftPos As Double
    Dim topPos As Double
    Dim widthVal As Double
    Dim cardWidth As Double
    Dim gapVal As Double

    leftPos = ws.Range("B5").Left
    topPos = ws.Range("B7").Top
    widthVal = ws.Range("B5:Q29").Width
    gapVal = 12
    cardWidth = (widthVal - (gapVal * 3)) / 4

    ws.Range("B6:M6").Merge
    ws.Range("B6").value = Dashboard_L("Synthèse exécutive", "Executive Summary")
    ws.Range("B6").Font.Bold = True
    ws.Range("B6").Font.Size = 15
    ws.Range("B6").Font.Color = RGB(20, 34, 51)

    With ws.Range("N6:Q6")
        .Merge
        .value = Dashboard_NoProjectLoadedText()
        .Font.Size = 9
        .Font.Color = RGB(96, 111, 128)
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlCenter
    End With

    Dashboard_UpdateKpiCard ws, 1, Dashboard_L("Avancement projet", "Project Progress"), Dashboard_NoDataText(), _
        Dashboard_NoProjectLoadedText(), _
        leftPos + 18, topPos, cardWidth - 10, 78, RGB(160, 170, 181)

    Dashboard_UpdateKpiCard ws, 2, Dashboard_L("Fin prévisionnelle", "Forecast Finish"), Dashboard_NoDataText(), _
        Dashboard_NoProjectLoadedText(), _
        leftPos + 18 + cardWidth + gapVal, topPos, cardWidth - 10, 78, RGB(160, 170, 181)

    Dashboard_UpdateKpiCard ws, 3, Dashboard_L("Activités critiques", "Critical Activities"), Dashboard_NoDataText(), _
        Dashboard_NoProjectLoadedText(), _
        leftPos + 18 + ((cardWidth + gapVal) * 2), topPos, cardWidth - 10, 78, RGB(160, 170, 181)

    Dashboard_UpdateKpiCard ws, 4, Dashboard_L("Momentum planning", "Schedule Momentum"), Dashboard_NoDataText(), _
        Dashboard_NoProjectLoadedText(), _
        leftPos + 18 + ((cardWidth + gapVal) * 3), topPos, cardWidth - 10, 78, RGB(160, 170, 181)

End Sub

'------------------------------------------------------------------------------
' FR: Actualise Render S Curve Empty sans modifier les regles metier qui produisent les donnees.
' EN: Refreshes Render S Curve Empty without changing the business rules that produce the data.
'------------------------------------------------------------------------------

Private Sub Dashboard_RenderSCurveEmpty(ByVal ws As Worksheet)

    Dim leftPos As Double
    Dim topPos As Double
    Dim widthVal As Double
    Dim heightVal As Double

    leftPos = ws.Range("B13").Left + 18
    topPos = ws.Range("B13").Top
    widthVal = (ws.Range("I13").Left + ws.Range("I13").Width) - leftPos
    heightVal = ws.Range("B13:I29").Height

    Dashboard_AddPanel ws, leftPos, topPos, widthVal, heightVal, RGB(255, 255, 255)
    Dashboard_AddCenteredSectionTitle ws, Dashboard_L("Snapshot S-Curve", "S-Curve Snapshot"), leftPos + 22, topPos + 10, widthVal - 44
    Dashboard_WriteDashboardEmptyState ws, leftPos + 22, topPos + 70, widthVal - 44

End Sub

'------------------------------------------------------------------------------
' FR: Actualise Render Planning Overview Empty sans modifier les regles metier qui produisent les donnees.
' EN: Refreshes Render Planning Overview Empty without changing the business rules that produce the data.
'------------------------------------------------------------------------------

Private Sub Dashboard_RenderPlanningOverviewEmpty(ByVal ws As Worksheet)

    Dim leftPos As Double
    Dim topPos As Double
    Dim widthVal As Double
    Dim heightVal As Double

    leftPos = ws.Range("J13").Left + 8
    topPos = ws.Range("J13").Top
    widthVal = ws.Range("J13:Q29").Width - 20
    heightVal = ws.Range("J13:Q29").Height

    Dashboard_AddPanel ws, leftPos, topPos, widthVal, heightVal, RGB(255, 255, 255)
    Dashboard_AddCenteredSectionTitle ws, Dashboard_L("Vue planning", "Planning Overview"), leftPos + 22, topPos + 10, widthVal - 44
    Dashboard_WriteDashboardEmptyState ws, leftPos + 22, topPos + 70, widthVal - 44

End Sub

'------------------------------------------------------------------------------
' FR: Actualise Render Hot Spots Empty sans modifier les regles metier qui produisent les donnees.
' EN: Refreshes Render Hot Spots Empty without changing the business rules that produce the data.
'------------------------------------------------------------------------------

Private Sub Dashboard_RenderHotSpotsEmpty(ByVal ws As Worksheet)

    Dim leftPos As Double
    Dim topPos As Double
    Dim widthVal As Double
    Dim heightVal As Double
    Dim cardTop As Double
    Dim cardHeight As Double
    Dim gapVal As Double
    Dim rightW As Double
    Dim largeW As Double
    Dim smallH As Double

    leftPos = ws.Range("B32").Left
    topPos = ws.Range("B32").Top
    widthVal = ws.Range("B32:Q56").Width
    heightVal = ws.Range("B32:Q56").Height

    Dashboard_AddSectionTitle ws, Dashboard_L("Points chauds", "Hot Spots"), leftPos + 22, topPos + 10, widthVal - 36

    cardTop = topPos + 48
    cardHeight = heightVal - 64
    gapVal = 14
    rightW = 240
    largeW = (widthVal - 44 - rightW - (gapVal * 2)) / 2
    smallH = (cardHeight - gapVal) / 2

    Dashboard_AddEmptyHotSpotCard ws, leftPos + 22, cardTop, largeW, cardHeight, Dashboard_L("Dérives majeures", "Top Delays")
    Dashboard_AddEmptyHotSpotCard ws, leftPos + 22 + largeW + gapVal, cardTop, largeW, cardHeight, Dashboard_L("Santé deadlines", "Deadline Health")
    Dashboard_AddEmptyHotSpotCard ws, leftPos + 22 + (largeW * 2) + (gapVal * 2), cardTop, rightW, smallH, Dashboard_L("Prochain jalon", "Next Milestone")
    Dashboard_AddEmptyHotSpotCard ws, leftPos + 22 + (largeW * 2) + (gapVal * 2), cardTop + smallH + gapVal, rightW, smallH, Dashboard_L("Activité critique", "Next Critical Activity")

End Sub
'------------------------------------------------------------------------------
' FR: Actualise Update Header Texts sans modifier les regles metier qui produisent les donnees.
' EN: Refreshes Update Header Texts without changing the business rules that produce the data.
' FR - Effet de bord : cree ou met a jour des shapes Excel.
' EN - Side effect: creates or updates Excel shapes.
'------------------------------------------------------------------------------

Private Sub Dashboard_UpdateHeaderTexts(ByVal ws As Worksheet)

    With ws.Range("B1:P2")
        .value = Dashboard_L("Tableau de bord", "Dashboard")
    End With

    With ws.Range("B3:P3")
        .value = Dashboard_L("Pilotage projet PM / Engineering - snapshot de comparaison", "PM / Engineering dashboard - comparison snapshots")
    End With

    On Error Resume Next
    ws.Shapes("btn_Update_Dashboard").TextFrame2.TextRange.Text = Dashboard_L("Nouveau snapshot", "New snapshot")
    ws.Shapes("btn_Dashboard_Refresh_Comparison").TextFrame2.TextRange.Text = Dashboard_L("Rafraîchir comparaison", "Refresh Comparison")
    ws.Shapes("btn_Dashboard_Reset").TextFrame2.TextRange.Text = Dashboard_L("Nettoyer Dashboard", "Clean Dashboard")
    On Error GoTo 0

End Sub

'------------------------------------------------------------------------------
' FR: Actualise Update Header Timestamp sans modifier les regles metier qui produisent les donnees.
' EN: Refreshes Update Header Timestamp without changing the business rules that produce the data.
'------------------------------------------------------------------------------

Private Sub Dashboard_UpdateHeaderTimestamp(ByVal ws As Worksheet)

    Dim shp As Shape
    Dim tsText As String
    Dim found As Boolean

    tsText = Dashboard_FormatTimestamp(Now)

    On Error Resume Next
    For Each shp In ws.Shapes
        If shp.TextFrame2.HasText Then
            If shp.Top < ws.Range("A4").Top Then
                If Dashboard_IsTimestampText(Trim$(shp.TextFrame2.TextRange.Text)) Then
                    shp.TextFrame2.TextRange.Text = tsText
                    found = True
                    Exit For
                End If
            End If
        End If
    Next shp
    On Error GoTo 0

    If Not found Then
        Dashboard_AddBadge ws, tsText, ws.Range("T1").Left, ws.Range("T1").Top + 5, 150, 22, RGB(226, 240, 255), RGB(29, 91, 158)
    End If

End Sub
'------------------------------------------------------------------------------
' FR: Reinitialise Remove Legacy Language Toggle dans le perimetre possede par le composant.
' EN: Resets Remove Legacy Language Toggle within the state owned by the component.
' FR - Effet de bord : efface uniquement les donnees ou objets cibles du contrat.
' EN - Side effect: clears only data or objects targeted by the contract.
'------------------------------------------------------------------------------

Private Sub Dashboard_RemoveLegacyLanguageToggle(ByVal ws As Worksheet)

    If ws Is Nothing Then Exit Sub

    On Error Resume Next
    ws.Shapes("DASH_Language_Label").Delete
    ws.Shapes("DASH_Language_BG").Delete
    ws.Shapes("DASH_Language_Knob").Delete
    On Error GoTo 0

End Sub

'------------------------------------------------------------------------------
' FR: Actualise Update Language Toggle sans modifier les regles metier qui produisent les donnees.
' EN: Refreshes Update Language Toggle without changing the business rules that produce the data.
' FR - Effet de bord : cree ou met a jour des shapes Excel.
' EN - Side effect: creates or updates Excel shapes.
'------------------------------------------------------------------------------

Private Sub Dashboard_UpdateLanguageToggle(ByVal ws As Worksheet)

    Dim trackShape As Shape
    Dim knobShape As Shape
    Dim isOn As Boolean
    Dim knobSize As Double
    Dim knobLeft As Double

    On Error Resume Next
    Set trackShape = ws.Shapes("DASH_Language_BG")
    Set knobShape = ws.Shapes("DASH_Language_Knob")
    On Error GoTo 0
    If trackShape Is Nothing Or knobShape Is Nothing Then Exit Sub

    isOn = Not Dashboard_IsFrench()
    Dashboard_FormatSwitchTrack trackShape, isOn

    knobSize = knobShape.Width
    If isOn Then
        knobLeft = trackShape.Left + trackShape.Width - knobSize - 2
    Else
        knobLeft = trackShape.Left + 2
    End If
    knobShape.Left = knobLeft
    knobShape.Top = trackShape.Top + ((trackShape.Height - knobSize) / 2)

End Sub

'------------------------------------------------------------------------------
' FR: Reinitialise Clear Kpi Shapes dans le perimetre possede par le composant.
' EN: Resets Clear Kpi Shapes within the state owned by the component.
' FR - Effet de bord : efface uniquement les donnees ou objets cibles du contrat.
' EN - Side effect: clears only data or objects targeted by the contract.
'------------------------------------------------------------------------------

Private Sub Dashboard_ClearKpiShapes(ByVal ws As Worksheet)

    Dim i As Long
    Dim shp As Shape
    Dim kpiLeft As Double
    Dim kpiRight As Double
    Dim kpiTop As Double
    Dim kpiBottom As Double

    kpiLeft = ws.Range("B7").Left - 4
    kpiRight = ws.Range("Q7").Left + ws.Range("Q7").Width + 4
    kpiTop = ws.Range("B7").Top - 6
    kpiBottom = ws.Range("B7").Top + 90

    For i = ws.Shapes.Count To 1 Step -1
        Set shp = ws.Shapes(i)
        If Left$(shp.Name, Len(DASH_PREFIX & "KPI")) = DASH_PREFIX & "KPI" Or _
            (shp.Left >= kpiLeft And shp.Left <= kpiRight And shp.Top >= kpiTop And shp.Top <= kpiBottom) Then
            shp.Delete
        End If
    Next i

End Sub

'------------------------------------------------------------------------------
' FR: Actualise Update Known Shape Texts sans modifier les regles metier qui produisent les donnees.
' EN: Refreshes Update Known Shape Texts without changing the business rules that produce the data.
'------------------------------------------------------------------------------

Private Sub Dashboard_UpdateKnownShapeTexts(ByVal ws As Worksheet)

    Dim shp As Shape
    Dim txt As String

    On Error Resume Next
    For Each shp In ws.Shapes
        If shp.TextFrame2.HasText Then
            txt = Trim$(shp.TextFrame2.TextRange.Text)
            Select Case txt
                Case "Snapshot S-Curve", "S-Curve Snapshot"
                    shp.TextFrame2.TextRange.Text = Dashboard_L("Snapshot S-Curve", "S-Curve Snapshot")
                Case "Vue planning", "Planning Overview"
                    shp.TextFrame2.TextRange.Text = Dashboard_L("Vue planning", "Planning Overview")
                Case "Points chauds", "Hot Spots"
                    shp.TextFrame2.TextRange.Text = Dashboard_L("Points chauds", "Hot Spots")
                Case "Dérives majeures", "Derives majeures", "Top Delays"
                    shp.TextFrame2.TextRange.Text = Dashboard_L("Dérives majeures", "Top Delays")
                Case "Risques jalons", "Deadline Risks"
                    shp.TextFrame2.TextRange.Text = Dashboard_L("Risques jalons", "Deadline Risks")
                Case "Alertes forecast", "Forecast Issues"
                    shp.TextFrame2.TextRange.Text = Dashboard_L("Alertes forecast", "Forecast Issues")
                Case "Aucune dérive planning détectée", "Aucune derive planning detectee", "No schedule delays detected"
                    shp.TextFrame2.TextRange.Text = Dashboard_L("Aucune dérive planning détectée", "No schedule delays detected")
                Case "Aucun risque deadline détecté", "Aucun risque deadline detecte", "No deadline risks detected"
                    shp.TextFrame2.TextRange.Text = Dashboard_L("Aucun risque deadline détecté", "No deadline risks detected")
                Case "Aucune alerte forecast détectée", "Aucune alerte forecast detectee", "No forecast issues detected"
                    shp.TextFrame2.TextRange.Text = Dashboard_L("Aucune alerte forecast détectée", "No forecast issues detected")
                Case "Aucun projet chargé", "Aucun projet charge", "No project loaded"
                    shp.TextFrame2.TextRange.Text = Dashboard_NoProjectLoadedText()
                Case "AUCUNE DONNÉE", "AUCUNE DONNEE", "NO DATA"
                    shp.TextFrame2.TextRange.Text = Dashboard_NoDataText()
            End Select
        End If
    Next shp
    On Error GoTo 0

End Sub

'------------------------------------------------------------------------------
' FR: Actualise Render Header sans modifier les regles metier qui produisent les donnees.
' EN: Refreshes Render Header without changing the business rules that produce the data.
'------------------------------------------------------------------------------

Private Sub Dashboard_RenderHeader(ByVal ws As Worksheet, Optional ByVal selectedFromLabel As String = "", Optional ByVal selectedToLabel As String = "")

    With ws.Range("B1:P2")
        .Merge
        .value = Dashboard_L("Tableau de bord", "Dashboard")
        .Font.Size = 24
        .Font.Bold = True
        .Font.Color = RGB(20, 34, 51)
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With

    With ws.Range("B3:P3")
        .Merge
        .value = Dashboard_L("Pilotage projet PM / Engineering - snapshot de comparaison", "PM / Engineering dashboard - comparison snapshots")
        .Font.Size = 10
        .Font.Color = RGB(96, 111, 128)
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With

    Dashboard_AddUpdateButton ws, ws.Range("Q1").Left, ws.Range("Q1").Top + 4, 126, 24
    Dashboard_AddBadge ws, Dashboard_FormatTimestamp(Now), ws.Range("T1").Left, ws.Range("T1").Top + 5, 150, 22, RGB(226, 240, 255), RGB(29, 91, 158)
    Dashboard_SetupSnapshotControls ws, selectedFromLabel, selectedToLabel, True
    Dashboard_AddResetButton ws, ws.Range("Q3").Left, ws.Range("Q3").Top + 4, 126, 24
    Dashboard_RemoveLegacyLanguageToggle ws

End Sub

'------------------------------------------------------------------------------
' FR: Ajoute la reference Update Button a la structure cible fournie par l'appelant.
' EN: Adds the Update Button reference to the target structure supplied by the caller.
' FR - Effet de bord : cree ou met a jour des shapes Excel.
' EN - Side effect: creates or updates Excel shapes.
'------------------------------------------------------------------------------

Private Sub Dashboard_AddUpdateButton(ByVal ws As Worksheet, ByVal x As Double, ByVal y As Double, ByVal w As Double, ByVal h As Double)

    Dim btn As Shape

    Set btn = ws.Shapes.AddShape(msoShapeRoundedRectangle, x, y, w, h)
    btn.Name = "btn_Update_Dashboard"
    btn.Fill.ForeColor.RGB = RGB(20, 34, 51)
    btn.Line.Visible = msoFalse
    btn.OnAction = "Run_Dashboard_Update"

    With btn.TextFrame2
        .MarginLeft = 8
        .MarginRight = 8
        .MarginTop = 2
        .MarginBottom = 2
        .TextRange.Text = Dashboard_L("Nouveau snapshot", "New snapshot")
        .TextRange.Font.Name = "Segoe UI"
        .TextRange.Font.Size = 9
        .TextRange.Font.Bold = msoTrue
        .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .VerticalAnchor = msoAnchorMiddle
    End With

End Sub

'------------------------------------------------------------------------------
' FR: Ajoute la reference Language Toggle a la structure cible fournie par l'appelant.
' EN: Adds the Language Toggle reference to the target structure supplied by the caller.
' FR - Effet de bord : cree ou met a jour des shapes Excel.
' EN - Side effect: creates or updates Excel shapes.
'------------------------------------------------------------------------------

Private Sub Dashboard_AddLanguageToggle(ByVal ws As Worksheet, ByVal x As Double, ByVal y As Double, ByVal w As Double, ByVal h As Double)

    Dim labelShape As Shape
    Dim trackShape As Shape
    Dim knobShape As Shape
    Dim labelW As Double
    Dim trackW As Double
    Dim trackH As Double
    Dim knobSize As Double
    Dim trackLeft As Double
    Dim knobLeft As Double
    Dim knobTop As Double
    Dim isOn As Boolean

    labelW = 56
    trackW = 36
    trackH = 14
    knobSize = 10
    trackLeft = x + labelW
    isOn = Not Dashboard_IsFrench()

    Set labelShape = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, x, y, labelW, trackH)
    labelShape.Name = "DASH_Language_Label"
    labelShape.OnAction = "Toggle_Dashboard_Language"
    labelShape.Placement = xlMove
    Dashboard_FormatSwitchLabel labelShape, "FR / EN"

    Set trackShape = ws.Shapes.AddShape(msoShapeRoundedRectangle, trackLeft, y, trackW, trackH)
    trackShape.Name = "DASH_Language_BG"
    trackShape.OnAction = "Toggle_Dashboard_Language"
    trackShape.Placement = xlMove
    trackShape.Adjustments.item(1) = 0.5
    Dashboard_FormatSwitchTrack trackShape, isOn

    knobTop = y + ((trackH - knobSize) / 2)
    If isOn Then
        knobLeft = trackLeft + trackW - knobSize - 2
    Else
        knobLeft = trackLeft + 2
    End If

    Set knobShape = ws.Shapes.AddShape(msoShapeOval, knobLeft, knobTop, knobSize, knobSize)
    knobShape.Name = "DASH_Language_Knob"
    knobShape.OnAction = "Toggle_Dashboard_Language"
    knobShape.Placement = xlMove
    Dashboard_FormatSwitchKnob knobShape

End Sub

'------------------------------------------------------------------------------
' FR: Normalise ou formate Format Switch Label selon le contrat canonique du composant.
' EN: Normalizes or formats Format Switch Label according to the component contract.
'------------------------------------------------------------------------------

Private Sub Dashboard_FormatSwitchLabel(ByVal shp As Shape, ByVal caption As String)

    With shp
        .Line.Visible = msoFalse
        .Fill.Visible = msoFalse
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .TextFrame2.MarginLeft = 0
        .TextFrame2.MarginRight = 0
        .TextFrame2.MarginTop = 0
        .TextFrame2.MarginBottom = 0
        .TextFrame2.TextRange.Text = caption
        .TextFrame2.TextRange.Font.Name = "Segoe UI"
        .TextFrame2.TextRange.Font.Size = 9
        .TextFrame2.TextRange.Font.Bold = msoTrue
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(68, 114, 196)
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    End With

End Sub

'------------------------------------------------------------------------------
' FR: Normalise ou formate Format Switch Track selon le contrat canonique du composant.
' EN: Normalizes or formats Format Switch Track according to the component contract.
'------------------------------------------------------------------------------

Private Sub Dashboard_FormatSwitchTrack(ByVal shp As Shape, ByVal isOn As Boolean)

    With shp
        .Line.Weight = 1
        .Shadow.Visible = msoFalse
        If isOn Then
            .Fill.ForeColor.RGB = RGB(68, 114, 196)
            .Line.ForeColor.RGB = RGB(68, 114, 196)
        Else
            .Fill.ForeColor.RGB = RGB(230, 230, 230)
            .Line.ForeColor.RGB = RGB(170, 170, 170)
        End If
    End With

End Sub

'------------------------------------------------------------------------------
' FR: Normalise ou formate Format Switch Knob selon le contrat canonique du composant.
' EN: Normalizes or formats Format Switch Knob according to the component contract.
'------------------------------------------------------------------------------

Private Sub Dashboard_FormatSwitchKnob(ByVal shp As Shape)

    With shp
        .Fill.ForeColor.RGB = RGB(255, 255, 255)
        .Line.ForeColor.RGB = RGB(150, 150, 150)
        .Line.Weight = 0.75
        .Shadow.Visible = msoFalse
    End With

End Sub

'------------------------------------------------------------------------------
' FR: Ajoute la reference Refresh Comparison Button a la structure cible fournie par l'appelant.
' EN: Adds the Refresh Comparison Button reference to the target structure supplied by the caller.
' FR - Effet de bord : cree ou met a jour des shapes Excel.
' EN - Side effect: creates or updates Excel shapes.
'------------------------------------------------------------------------------

Private Sub Dashboard_AddRefreshComparisonButton(ByVal ws As Worksheet, ByVal x As Double, ByVal y As Double, ByVal w As Double, ByVal h As Double)

    Dim btn As Shape

    Set btn = ws.Shapes.AddShape(msoShapeRoundedRectangle, x, y, w, h)
    btn.Name = "btn_Dashboard_Refresh_Comparison"
    btn.Fill.ForeColor.RGB = RGB(96, 111, 128)
    btn.Line.Visible = msoFalse
    btn.OnAction = "Refresh_Dashboard_Comparison"

    With btn.TextFrame2
        .MarginLeft = 8
        .MarginRight = 8
        .MarginTop = 2
        .MarginBottom = 2
        .TextRange.Text = Dashboard_L("Rafraîchir comparaison", "Refresh Comparison")
        .TextRange.Font.Name = "Segoe UI"
        .TextRange.Font.Size = 9
        .TextRange.Font.Bold = msoTrue
        .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .VerticalAnchor = msoAnchorMiddle
    End With

End Sub

'------------------------------------------------------------------------------
' FR: Ajoute la reference Reset Button a la structure cible fournie par l'appelant.
' EN: Adds the Reset Button reference to the target structure supplied by the caller.
' FR - Effet de bord : cree ou met a jour des shapes Excel.
' EN - Side effect: creates or updates Excel shapes.
'------------------------------------------------------------------------------

Private Sub Dashboard_AddResetButton(ByVal ws As Worksheet, ByVal x As Double, ByVal y As Double, ByVal w As Double, ByVal h As Double)

    Dim btn As Shape

    Set btn = ws.Shapes.AddShape(msoShapeRoundedRectangle, x, y, w, h)
    btn.Name = "btn_Dashboard_Reset"
    btn.Fill.ForeColor.RGB = RGB(192, 80, 77)
    btn.Line.Visible = msoFalse
    btn.OnAction = "Reset_Dashboard"

    With btn.TextFrame2
        .MarginLeft = 8
        .MarginRight = 8
        .MarginTop = 2
        .MarginBottom = 2
        .TextRange.Text = Dashboard_L("Nettoyer Dashboard", "Clean Dashboard")
        .TextRange.Font.Name = "Segoe UI"
        .TextRange.Font.Size = 9
        .TextRange.Font.Bold = msoTrue
        .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .VerticalAnchor = msoAnchorMiddle
    End With

End Sub

'------------------------------------------------------------------------------
' FR: Cree ou remet en conformite la reference Reset Button de maniere idempotente.
' EN: Creates or restores the Reset Button reference to a compliant state idempotently.
'------------------------------------------------------------------------------

Private Sub Dashboard_EnsureResetButton(ByVal ws As Worksheet)

    Dim btn As Shape

    Set btn = Dashboard_GetShape(ws, "btn_Dashboard_Reset")
    If btn Is Nothing Then
        Dashboard_AddResetButton ws, ws.Range("Q3").Left, ws.Range("Q3").Top + 4, 126, 24
    Else
        Dashboard_PositionKpiShape btn, ws.Range("Q3").Left, ws.Range("Q3").Top + 4, 126, 24
        btn.Fill.ForeColor.RGB = RGB(192, 80, 77)
        btn.Line.Visible = msoFalse
        btn.Visible = msoTrue
        btn.OnAction = "Reset_Dashboard"
        btn.TextFrame2.TextRange.Text = Dashboard_L("Nettoyer Dashboard", "Clean Dashboard")
    End If

End Sub
'------------------------------------------------------------------------------
' FR: Active ou initialise Setup Snapshot Controls dans l'etat runtime du composant.
' EN: Activates or initializes Setup Snapshot Controls in the component runtime state.
' FR - Effet de bord : efface uniquement les donnees ou objets cibles du contrat.
' EN - Side effect: clears only data or objects targeted by the contract.
'------------------------------------------------------------------------------

Private Sub Dashboard_SetupSnapshotControls(ByVal ws As Worksheet, ByVal selectedFromLabel As String, ByVal selectedToLabel As String, Optional ByVal recreateButton As Boolean = True)

    Dim tbl As ListObject
    Dim fromDefault As String
    Dim toDefault As String

    Set tbl = Dashboard_EnsureSnapshotsTable()
    Dashboard_GetDefaultSnapshotLabels tbl, fromDefault, toDefault

    If selectedFromLabel <> "" And Dashboard_SnapshotLabelExists(tbl, selectedFromLabel) Then fromDefault = selectedFromLabel
    If selectedToLabel <> "" And Dashboard_SnapshotLabelExists(tbl, selectedToLabel) Then toDefault = selectedToLabel

    ws.Range("B4").value = Dashboard_L("De", "From")
    ws.Range("F4").value = Dashboard_L("A", "To")
    ws.Range("B4,F4").Font.Bold = True
    ws.Range("B4,F4").Font.Color = RGB(96, 111, 128)

    ws.Range("C4:D4").Merge
    ws.Range("G4:H4").Merge

    With ws.Range("C4:D4")
        .value = fromDefault
        .Interior.Color = RGB(255, 255, 255)
        .Font.Color = RGB(20, 34, 51)
        .HorizontalAlignment = xlLeft
        .Validation.Delete
        If Not tbl.DataBodyRange Is Nothing Then .Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="=DashboardSnapshotLabels"
    End With

    With ws.Range("G4:H4")
        .value = toDefault
        .Interior.Color = RGB(255, 255, 255)
        .Font.Color = RGB(20, 34, 51)
        .HorizontalAlignment = xlLeft
        .Validation.Delete
        If Not tbl.DataBodyRange Is Nothing Then .Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="=DashboardSnapshotLabels"
    End With

    If recreateButton Then Dashboard_AddRefreshComparisonButton ws, ws.Range("J4").Left, ws.Range("J4").Top - 1, 132, 22

End Sub

'------------------------------------------------------------------------------
' FR: Actualise Render Executive Summary sans modifier les regles metier qui produisent les donnees.
' EN: Refreshes Render Executive Summary without changing the business rules that produce the data.
'------------------------------------------------------------------------------

Private Sub Dashboard_RenderExecutiveSummary( _
    ByVal ws As Worksheet, _
    ByVal tblWBS As ListObject, _
    ByVal mapWBS As Object, _
    ByVal tblCalc As ListObject, _
    ByVal mapCalc As Object, _
    ByVal scurveProjection As Object)

    Dashboard_UpdateKpiCardsInPlace ws, tblWBS, mapWBS, tblCalc, mapCalc, scurveProjection

End Sub

'------------------------------------------------------------------------------
' FR: Actualise Update Kpi Cards In Place sans modifier les regles metier qui produisent les donnees.
' EN: Refreshes Update Kpi Cards In Place without changing the business rules that produce the data.
'------------------------------------------------------------------------------

Private Sub Dashboard_UpdateKpiCardsInPlace( _
    ByVal ws As Worksheet, _
    ByVal tblWBS As ListObject, _
    ByVal mapWBS As Object, _
    ByVal tblCalc As ListObject, _
    ByVal mapCalc As Object, _
    ByVal scurveProjection As Object)

    Dim actualProgress As Double
    Dim plannedProgress As Double
    Dim progressVar As Double
    Dim baselineFinish As Variant
    Dim calcFinish As Variant
    Dim driftDays As Variant
    Dim deadlineExceeded As Long
    Dim forecastIssues As Long
    Dim cpCount As Long
    Dim lpCount As Long
    Dim lpRemaining As Long
    Dim statusText As String
    Dim statusColor As Long
    Dim contractText As String
    Dim contractColor As Long
    Dim momentumText As String
    Dim momentumColor As Long
    Dim compareText As String
    Dim momentumSubText As String
    Dim progressDeltaText As String
    Dim forecastDeltaText As String
    Dim progressColor As Long
    Dim forecastColor As Long
    Dim criticalHeroText As String
    Dim criticalSubText As String
    Dim criticalColor As Long
    Dim fromRow As Object
    Dim toRow As Object
    Dim fromLabel As String
    Dim toLabel As String
    Dim leftPos As Double
    Dim topPos As Double
    Dim widthVal As Double
    Dim cardWidth As Double
    Dim gapVal As Double

    If tblCalc Is Nothing Or tblCalc.DataBodyRange Is Nothing Then
        Dashboard_RenderExecutiveSummaryEmpty ws
        Exit Sub
    End If

    actualProgress = CDbl(scurveProjection("ActualProgress"))
    plannedProgress = CDbl(scurveProjection("PlannedProgress"))
    progressVar = actualProgress - plannedProgress

    baselineFinish = Dashboard_MaxDate(tblCalc, mapCalc, "Baseline Finish")
    calcFinish = Dashboard_MaxDate(tblCalc, mapCalc, "Calculated Finish")
    If Dashboard_HasDateValue(baselineFinish) And Dashboard_HasDateValue(calcFinish) Then driftDays = CLng(Dashboard_DateNumber(calcFinish) - Dashboard_DateNumber(baselineFinish))

    deadlineExceeded = Dashboard_CountNumericBelow(tblCalc, mapCalc, "Deadline Float", 0#)
    forecastIssues = Dashboard_CountErrorContains(tblCalc, mapCalc, "Forecast")
    cpCount = Dashboard_CountMarker(tblCalc, mapCalc, "Critical Path")
    lpCount = Dashboard_CountMarker(tblCalc, mapCalc, "Longest Path")
    lpRemaining = Dashboard_CountMarkerRemaining(tblCalc, mapCalc, "Longest Path")

    If deadlineExceeded > 0 Or forecastIssues > 0 Then
        statusText = "ACTION"
        statusColor = RGB(192, 80, 77)
    ElseIf HasValue(driftDays) And CLng(driftDays) > 0 Then
        statusText = "WATCH"
        statusColor = RGB(238, 156, 68)
    Else
        statusText = "ON TRACK"
        statusColor = RGB(0, 145, 112)
    End If

    fromLabel = CStr(ws.Range(DASH_FROM_CELL).value)
    toLabel = CStr(ws.Range(DASH_TO_CELL).value)
    Dashboard_GetComparisonRows fromLabel, toLabel, fromRow, toRow

    If Not toRow Is Nothing Then
        actualProgress = CDbl(toRow("ActualProgress"))
        plannedProgress = CDbl(toRow("PlannedProgress"))
        progressVar = CDbl(toRow("ProgressVariance"))
        calcFinish = toRow("ForecastFinish")
        baselineFinish = toRow("BaselineFinish")
        driftDays = toRow("DriftDays")
        contractText = CStr(toRow("ContractStatus"))
    Else
        contractText = Dashboard_ContractStatus(driftDays, contractColor)
    End If

    contractText = Dashboard_ContractStatus(driftDays, contractColor)
    compareText = Dashboard_ComparisonText(fromRow, toRow)
    progressDeltaText = Dashboard_BehindPlanText(plannedProgress, actualProgress, progressColor)
    forecastDeltaText = Dashboard_ContractDriftText(driftDays, forecastColor)
    criticalHeroText = Dashboard_CriticalActivitiesHero(fromRow, toRow, criticalColor, criticalSubText)
    momentumText = Dashboard_ScheduleMomentumStatus(fromRow, toRow, momentumColor, momentumSubText)

    leftPos = ws.Range("B5").Left
    topPos = ws.Range("B5").Top
    widthVal = ws.Range("B5:Q29").Width
    gapVal = 12
    cardWidth = (widthVal - (gapVal * 3)) / 4

    ws.Range("B6:M6").Merge
    ws.Range("B6").value = Dashboard_L("Synthèse exécutive", "Executive Summary")
    ws.Range("B6").Font.Bold = True
    ws.Range("B6").Font.Size = 15
    ws.Range("B6").Font.Color = RGB(20, 34, 51)

    Dashboard_UpdateKpiCard ws, 1, Dashboard_L("Avancement projet", "Project Progress"), Format$(actualProgress, "0%"), _
        progressDeltaText, _
        leftPos + 18, ws.Range("B7").Top, cardWidth - 10, 78, progressColor

    Dashboard_UpdateKpiCard ws, 2, Dashboard_L("Fin prévisionnelle", "Forecast Finish"), Dashboard_FormatDateShort(calcFinish), _
        forecastDeltaText, _
        leftPos + 18 + cardWidth + gapVal, ws.Range("B7").Top, cardWidth - 10, 78, forecastColor

    Dashboard_UpdateKpiCard ws, 3, Dashboard_L("Activités critiques", "Critical Activities"), criticalHeroText, _
        criticalSubText, _
        leftPos + 18 + ((cardWidth + gapVal) * 2), ws.Range("B7").Top, cardWidth - 10, 78, criticalColor

    Dashboard_UpdateKpiCard ws, 4, Dashboard_L("Momentum planning", "Schedule Momentum"), Dashboard_LocalMomentumStatus(momentumText), _
        momentumSubText, _
        leftPos + 18 + ((cardWidth + gapVal) * 3), ws.Range("B7").Top, cardWidth - 10, 78, momentumColor

    With ws.Range("N6:Q6")
        .Merge
        .value = compareText
        .Font.Size = 9
        .Font.Color = RGB(96, 111, 128)
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlCenter
    End With

End Sub

'------------------------------------------------------------------------------
' FR: Actualise Update S Curve Chart In Place sans modifier les regles metier qui produisent les donnees.
' EN: Refreshes Update S Curve Chart In Place without changing the business rules that produce the data.
' FR - Effet de bord : efface uniquement les donnees ou objets cibles du contrat.
' EN - Side effect: clears only data or objects targeted by the contract.
'------------------------------------------------------------------------------

Private Sub Dashboard_UpdateSCurveChartInPlace(ByVal ws As Worksheet, ByVal scurveProjection As Object)

    Dim chartObj As ChartObject
    Dim ch As Chart
    Dim leftPos As Double
    Dim topPos As Double
    Dim widthVal As Double
    Dim heightVal As Double
    Dim xRange As Range
    Dim baselineRange As Range
    Dim actualRange As Range
    Dim forecastRange As Range

    If scurveProjection Is Nothing Then Exit Sub
    If Not scurveProjection.Exists("DateRange") Then Exit Sub
    Set xRange = scurveProjection("DateRange")
    If scurveProjection.Exists("BaselineRange") Then Set baselineRange = scurveProjection("BaselineRange")
    If scurveProjection.Exists("ActualRange") Then Set actualRange = scurveProjection("ActualRange")
    If scurveProjection.Exists("ForecastRange") Then Set forecastRange = scurveProjection("ForecastRange")

    On Error Resume Next
    Set chartObj = ws.ChartObjects(DASH_CHART_SCURVE)
    On Error GoTo 0
    If chartObj Is Nothing Then
        Dashboard_RenderSCurveChart ws, scurveProjection
        Exit Sub
    End If

    leftPos = ws.Range("B13").Left + 18
    topPos = ws.Range("B13").Top
    widthVal = (ws.Range("I13").Left + ws.Range("I13").Width) - leftPos
    heightVal = ws.Range("B13:I29").Height
    chartObj.Left = leftPos + 16
    chartObj.Top = topPos + 38
    chartObj.Width = widthVal - 24
    chartObj.Height = heightVal - 56

    Set ch = chartObj.Chart
    Do While ch.SeriesCollection.Count > 0
        ch.SeriesCollection(1).Delete
    Loop
    ch.ChartType = xlLine
    ch.HasTitle = False
    ch.HasLegend = True
    ch.Legend.Position = xlLegendPositionBottom
    Dashboard_AddLineSeries ch, Dashboard_L("Référence", "Baseline"), xRange, baselineRange, RGB(150, 150, 150), 1.5, False
    Dashboard_AddLineSeries ch, Dashboard_L("Réel", "Actual"), xRange, actualRange, RGB(0, 145, 112), 2.5, False
    Dashboard_AddLineSeries ch, Dashboard_L("Prévision", "Forecast"), xRange, forecastRange, RGB(43, 106, 176), 2.5, True

    On Error Resume Next
    ch.Axes(xlValue).TickLabels.NumberFormat = "0%"
    ch.Axes(xlCategory).TickLabels.NumberFormat = Dashboard_ChartDateNumberFormat(False)
    ch.Axes(xlCategory).TickLabels.Orientation = 45
    ch.Axes(xlValue).MajorGridlines.Format.Line.ForeColor.RGB = RGB(225, 231, 238)
    ch.ChartArea.Format.Line.Visible = msoFalse
    ch.PlotArea.Format.Line.Visible = msoFalse
    ch.ChartArea.Format.Fill.Visible = msoFalse
    ch.PlotArea.Format.Fill.Visible = msoFalse
    On Error GoTo 0
    SCurve_DrawDashboardTodayLine ch, ws, DASH_SCURVE_TODAY_LINE

End Sub
'------------------------------------------------------------------------------
' FR: Actualise Render S Curve Chart sans modifier les regles metier qui produisent les donnees.
' EN: Refreshes Render S Curve Chart without changing the business rules that produce the data.
'------------------------------------------------------------------------------

Private Sub Dashboard_RenderSCurveChart(ByVal ws As Worksheet, ByVal scurveProjection As Object)

    Dim chartObj As ChartObject
    Dim ch As Chart
    Dim xRange As Range
    Dim baselineRange As Range
    Dim actualRange As Range
    Dim forecastRange As Range
    Dim leftPos As Double
    Dim topPos As Double
    Dim widthVal As Double
    Dim heightVal As Double

    If scurveProjection Is Nothing Then Exit Sub
    If Not scurveProjection.Exists("DateRange") Then Exit Sub
    Set xRange = scurveProjection("DateRange")
    If scurveProjection.Exists("BaselineRange") Then Set baselineRange = scurveProjection("BaselineRange")
    If scurveProjection.Exists("ActualRange") Then Set actualRange = scurveProjection("ActualRange")
    If scurveProjection.Exists("ForecastRange") Then Set forecastRange = scurveProjection("ForecastRange")
    leftPos = ws.Range("B13").Left + 18
    topPos = ws.Range("B13").Top
    widthVal = (ws.Range("I13").Left + ws.Range("I13").Width) - leftPos
    heightVal = ws.Range("B13:I29").Height
    Dashboard_AddPanel ws, leftPos, topPos, widthVal, heightVal, RGB(255, 255, 255)
    Dashboard_AddCenteredSectionTitle ws, Dashboard_L("Snapshot S-Curve", "S-Curve Snapshot"), leftPos + 22, topPos + 10, widthVal - 44
    Set chartObj = ws.ChartObjects.Add(leftPos + 16, topPos + 38, widthVal - 24, heightVal - 56)
    chartObj.Name = DASH_CHART_SCURVE
    Set ch = chartObj.Chart
    ch.ChartType = xlLine
    ch.HasTitle = False
    ch.HasLegend = True
    ch.Legend.Position = xlLegendPositionBottom
    Dashboard_AddLineSeries ch, Dashboard_L("Référence", "Baseline"), xRange, baselineRange, RGB(150, 150, 150), 1.5, False
    Dashboard_AddLineSeries ch, Dashboard_L("Réel", "Actual"), xRange, actualRange, RGB(0, 145, 112), 2.5, False
    Dashboard_AddLineSeries ch, Dashboard_L("Prévision", "Forecast"), xRange, forecastRange, RGB(43, 106, 176), 2.5, True
    On Error Resume Next
    ch.Axes(xlValue).TickLabels.NumberFormat = "0%"
    ch.Axes(xlCategory).TickLabels.NumberFormat = Dashboard_ChartDateNumberFormat(False)
    ch.Axes(xlCategory).TickLabels.Orientation = 45
    ch.Axes(xlValue).MajorGridlines.Format.Line.ForeColor.RGB = RGB(225, 231, 238)
    ch.ChartArea.Format.Line.Visible = msoFalse
    ch.PlotArea.Format.Line.Visible = msoFalse
    ch.ChartArea.Format.Fill.Visible = msoFalse
    ch.PlotArea.Format.Fill.Visible = msoFalse
    On Error GoTo 0
    SCurve_DrawDashboardTodayLine ch, ws, DASH_SCURVE_TODAY_LINE

End Sub

'------------------------------------------------------------------------------
' FR: Ajoute la reference Line Series a la structure cible fournie par l'appelant.
' EN: Adds the Line Series reference to the target structure supplied by the caller.
'------------------------------------------------------------------------------

Private Sub Dashboard_AddLineSeries( _
    ByVal ch As Chart, _
    ByVal seriesName As String, _
    ByVal xRange As Range, _
    ByVal yRange As Range, _
    ByVal rgbColor As Long, _
    ByVal weightVal As Double, _
    ByVal dashed As Boolean)

    Dim s As Series
    If xRange Is Nothing Or yRange Is Nothing Then Exit Sub
    Set s = ch.SeriesCollection.NewSeries
    With s
        .Name = seriesName
        .XValues = xRange
        .Values = yRange
        .ChartType = xlLine
        .Format.Line.ForeColor.RGB = rgbColor
        .Format.Line.Weight = weightVal
        If dashed Then .Format.Line.DashStyle = msoLineDash
        .MarkerStyle = xlMarkerStyleNone
    End With

End Sub

'------------------------------------------------------------------------------
' FR: Actualise Update Planning Overview Content sans modifier les regles metier qui produisent les donnees.
' EN: Refreshes Update Planning Overview Content without changing the business rules that produce the data.
'------------------------------------------------------------------------------

Private Sub Dashboard_UpdatePlanningOverviewContent(ByVal ws As Worksheet, ByVal tblWBS As ListObject, ByVal mapWBS As Object, ByVal tblCalc As ListObject, ByVal mapCalc As Object)

    Dim leftPos As Double
    Dim topPos As Double
    Dim widthVal As Double
    Dim heightVal As Double

    leftPos = ws.Range("J13").Left + 8
    topPos = ws.Range("J13").Top
    widthVal = ws.Range("J13:Q29").Width - 20
    heightVal = ws.Range("J13:Q29").Height

    Dashboard_DeleteShapesInRect ws, leftPos + 1, topPos + 36, widthVal - 2, heightVal - 38
    Dashboard_RenderPlanningOverview ws, tblWBS, mapWBS, tblCalc, mapCalc, False

End Sub
'------------------------------------------------------------------------------
' FR: Actualise Render Planning Overview sans modifier les regles metier qui produisent les donnees.
' EN: Refreshes Render Planning Overview without changing the business rules that produce the data.
'------------------------------------------------------------------------------

Private Sub Dashboard_RenderPlanningOverview(ByVal ws As Worksheet, ByVal tblWBS As ListObject, ByVal mapWBS As Object, ByVal tblCalc As ListObject, ByVal mapCalc As Object, Optional ByVal renderShell As Boolean = True)

    Dim leftPos As Double
    Dim topPos As Double
    Dim widthVal As Double
    Dim heightVal As Double
    Dim arr As Variant
    Dim r As Long
    Dim rowsToShow As Collection
    Dim rowRef As Variant
    Dim rowIndex As Long
    Dim rowCount As Long
    Dim shown As Long
    Dim rowH As Double
    Dim rowCenter As Double
    Dim shapeCenterY As Double
    Dim separatorTopY As Double
    Dim separatorBottomY As Double
    Dim firstRowCenter As Double
    Dim labelTop As Double
    Dim labelText As String
    Dim rowStart As Variant
    Dim rowFinish As Variant
    Dim baseStart As Variant
    Dim baseFinish As Variant
    Dim x1 As Double
    Dim x2 As Double
    Dim todayX As Double
    Dim axisLeft As Double
    Dim axisWidth As Double
    Dim axisRight As Double
    Dim progressSummary As Double
    Dim progressById As Object
    Dim axisMin As Double
    Dim axisMax As Double
    Dim axisMargin As Double
    Dim usefulLeft As Double
    Dim usefulTop As Double
    Dim usefulWidth As Double
    Dim usefulHeight As Double
    Dim labelWidth As Double
    Dim plotTop As Double
    Dim plotBottom As Double
    Dim axisY As Double
    Dim fontSize As Double
    Dim barHeight As Double
    Dim baselineHeight As Double
    Dim wbsVal As String
    Dim taskName As String
    Dim isMilestone As Boolean

    If tblCalc Is Nothing Then Exit Sub
    If tblCalc.DataBodyRange Is Nothing Then Exit Sub
    If Not mapCalc.Exists("IsSummary") Then Exit Sub
    If Not mapCalc.Exists("Calculated Start") Then Exit Sub
    If Not mapCalc.Exists("Calculated Finish") Then Exit Sub

    leftPos = ws.Range("J13").Left + 8
    topPos = ws.Range("J13").Top
    widthVal = ws.Range("J13:Q29").Width - 20
    heightVal = ws.Range("J13:Q29").Height

    If renderShell Then
        Dashboard_AddPanel ws, leftPos, topPos, widthVal, heightVal, RGB(255, 255, 255)
        Dashboard_AddCenteredSectionTitle ws, Dashboard_L("Vue planning", "Planning Overview"), leftPos + 22, topPos + 10, widthVal - 44
    End If

    arr = tblCalc.DataBodyRange.value
    Set progressById = Dashboard_BuildProgressById(tblWBS, mapWBS)

    Set rowsToShow = Dashboard_PlanningOverviewRows(arr, mapCalc, axisMin, axisMax)
    If rowsToShow.Count = 0 Or axisMin <= 0 Or axisMax < axisMin Then
        Dashboard_WriteEmptyState ws, leftPos + 22, topPos + 70, widthVal - 44, Dashboard_L("Aucune donnée planning summary disponible", "No summary schedule data available")
        Exit Sub
    End If

    axisMargin = Dashboard_PlanningAxisMargin(axisMin, axisMax)
    axisMin = axisMin - axisMargin
    axisMax = axisMax + axisMargin
    If axisMax <= axisMin Then axisMax = axisMin + 1

    usefulLeft = leftPos + 22
    usefulTop = topPos + 42
    usefulWidth = widthVal - 44
    usefulHeight = heightVal - 62
    labelWidth = WorksheetFunction.Min(usefulWidth * 0.25, 132)
    axisLeft = usefulLeft + labelWidth + 14
    axisWidth = usefulWidth - labelWidth - 14
    axisRight = axisLeft + axisWidth
    axisY = usefulTop + 18
    plotTop = usefulTop + 46
    plotBottom = topPos + heightVal - 36

    Dashboard_RenderTimeAxis ws, axisMin, axisMax, axisLeft, axisY, axisWidth

    rowCount = rowsToShow.Count
    If rowCount <= 0 Then Exit Sub
    rowH = (plotBottom - plotTop) / rowCount
    If rowH > 22 Then rowH = 22
    If rowH < 6 Then rowH = 6

    fontSize = rowH - 2
    If fontSize > 9 Then fontSize = 9
    If fontSize < 8 Then fontSize = 8

    barHeight = rowH * 0.42
    If barHeight > 8 Then barHeight = 8
    If barHeight < 4 Then barHeight = 4
    baselineHeight = 3

    firstRowCenter = plotTop + ((plotBottom - plotTop) - (rowH * rowCount)) / 2 + (rowH / 2)
    If firstRowCenter < plotTop + (rowH / 2) Then firstRowCenter = plotTop + (rowH / 2)

    shown = 0
    For Each rowRef In rowsToShow
        rowIndex = CLng(rowRef)
        rowStart = GetCellValue(Dashboard_ArrayVal(arr, mapCalc, rowIndex, "Calculated Start"))
        rowFinish = GetCellValue(Dashboard_ArrayVal(arr, mapCalc, rowIndex, "Calculated Finish"))
        If Dashboard_HasDateValue(rowStart) And Dashboard_HasDateValue(rowFinish) Then
            shown = shown + 1
            rowCenter = firstRowCenter + ((shown - 1) * rowH)
            shapeCenterY = rowCenter + 1.5
            separatorTopY = shapeCenterY - (rowH / 2)
            separatorBottomY = shapeCenterY + (rowH / 2)

            If shown = 1 Then Dashboard_AddLine ws, usefulLeft, separatorTopY, axisRight, separatorTopY, RGB(224, 230, 237), 0.5
            Dashboard_AddLine ws, usefulLeft, separatorBottomY, axisRight, separatorBottomY, RGB(224, 230, 237), 0.5

            wbsVal = CStr(Dashboard_ArrayVal(arr, mapCalc, rowIndex, "WBS"))
            taskName = CStr(Dashboard_ArrayVal(arr, mapCalc, rowIndex, "Task Name"))
            labelText = Dashboard_TruncateText(wbsVal & "  " & taskName, 26)
            labelTop = shapeCenterY - ((fontSize + 4) / 2)
            Dashboard_AddMiniGanttLabel ws, labelText, usefulLeft, labelTop, labelWidth, fontSize

            progressSummary = Dashboard_SummaryProgress(arr, mapCalc, progressById, rowIndex)
            isMilestone = (Dashboard_DateNumber(rowFinish) <= Dashboard_DateNumber(rowStart))

            baseStart = GetCellValue(Dashboard_ArrayVal(arr, mapCalc, rowIndex, "Baseline Start"))
            baseFinish = GetCellValue(Dashboard_ArrayVal(arr, mapCalc, rowIndex, "Baseline Finish"))
            If Dashboard_HasDateValue(baseStart) And Dashboard_HasDateValue(baseFinish) Then
                x1 = Dashboard_DateToX(Dashboard_DateNumber(baseStart), axisMin, axisMax, axisLeft, axisWidth)
                x2 = Dashboard_DateToX(Dashboard_DateNumber(baseFinish), axisMin, axisMax, axisLeft, axisWidth)
                If (x2 - x1) >= 6 Then
                    Dashboard_AddBar ws, x1, shapeCenterY - (baselineHeight / 2), x2 - x1, baselineHeight, RGB(189, 198, 208)
                End If
            End If

            x1 = Dashboard_DateToX(Dashboard_DateNumber(rowStart), axisMin, axisMax, axisLeft, axisWidth)
            x2 = Dashboard_DateToX(Dashboard_DateNumber(rowFinish), axisMin, axisMax, axisLeft, axisWidth)
            If isMilestone Then
                Dashboard_AddMilestoneDiamond ws, x1, shapeCenterY, WorksheetFunction.Max(7, barHeight + 2), Dashboard_MiniGanttCurrentColor(progressSummary)
            ElseIf (x2 - x1) >= 6 Then
                Dashboard_AddProgressBar ws, x1, shapeCenterY - (barHeight / 2), x2 - x1, barHeight, progressSummary, rowStart, rowFinish
            End If
        End If
    Next rowRef

    If shown = 0 Then
        Dashboard_WriteEmptyState ws, leftPos + 22, topPos + 70, widthVal - 44, Dashboard_L("Aucune donnée planning summary disponible", "No summary schedule data available")
    End If

    If CDbl(Date) >= axisMin And CDbl(Date) <= axisMax Then
        todayX = Dashboard_DateToX(CDbl(Date), axisMin, axisMax, axisLeft, axisWidth)
        Dashboard_AddLine ws, todayX, plotTop, todayX, plotBottom, RGB(0, 145, 112), 1.25
    End If

End Sub

'------------------------------------------------------------------------------
' FR: Actualise Update Hot Spots Content sans modifier les regles metier qui produisent les donnees.
' EN: Refreshes Update Hot Spots Content without changing the business rules that produce the data.
'------------------------------------------------------------------------------

Private Sub Dashboard_UpdateHotSpotsContent(ByVal ws As Worksheet, ByVal tblWBS As ListObject, ByVal mapWBS As Object, ByVal tblCalc As ListObject, ByVal mapCalc As Object)

    Dim leftPos As Double
    Dim topPos As Double
    Dim widthVal As Double
    Dim heightVal As Double

    leftPos = ws.Range("B32").Left
    topPos = ws.Range("B32").Top
    widthVal = ws.Range("B32:Q56").Width
    heightVal = ws.Range("B32:Q56").Height

    Dashboard_DeleteShapesInRect ws, leftPos + 1, topPos + 40, widthVal - 2, heightVal - 42
    Dashboard_RenderHotSpots ws, tblWBS, mapWBS, tblCalc, mapCalc, False

End Sub
'------------------------------------------------------------------------------
' FR: Actualise Render Hot Spots sans modifier les regles metier qui produisent les donnees.
' EN: Refreshes Render Hot Spots without changing the business rules that produce the data.
'------------------------------------------------------------------------------

Private Sub Dashboard_RenderHotSpots(ByVal ws As Worksheet, ByVal tblWBS As ListObject, ByVal mapWBS As Object, ByVal tblCalc As ListObject, ByVal mapCalc As Object, Optional ByVal renderShell As Boolean = True)

    Dim leftPos As Double
    Dim topPos As Double
    Dim widthVal As Double
    Dim heightVal As Double
    Dim cardTop As Double
    Dim cardHeight As Double
    Dim gapVal As Double
    Dim rightW As Double
    Dim largeW As Double
    Dim smallH As Double

    leftPos = ws.Range("B32").Left
    topPos = ws.Range("B32").Top
    widthVal = ws.Range("B32:Q56").Width
    heightVal = ws.Range("B32:Q56").Height

    If renderShell Then Dashboard_AddSectionTitle ws, Dashboard_L("Points chauds", "Hot Spots"), leftPos + 22, topPos + 10, widthVal - 36

    cardTop = topPos + 48
    cardHeight = heightVal - 64
    gapVal = 14
    rightW = 240
    largeW = (widthVal - 44 - rightW - (gapVal * 2)) / 2
    smallH = (cardHeight - gapVal) / 2

    Dashboard_RenderTopDelays ws, tblWBS, mapWBS, leftPos + 22, cardTop, largeW, cardHeight
    Dashboard_RenderDeadlineHealth ws, tblCalc, mapCalc, leftPos + 22 + largeW + gapVal, cardTop, largeW, cardHeight
    Dashboard_RenderNextMilestone ws, tblCalc, mapCalc, leftPos + 22 + (largeW * 2) + (gapVal * 2), cardTop, rightW, smallH
    Dashboard_RenderNextCriticalActivity ws, tblCalc, mapCalc, leftPos + 22 + (largeW * 2) + (gapVal * 2), cardTop + smallH + gapVal, rightW, smallH

End Sub

'------------------------------------------------------------------------------
' FR: Actualise Render Top Delays sans modifier les regles metier qui produisent les donnees.
' EN: Refreshes Render Top Delays without changing the business rules that produce the data.
'------------------------------------------------------------------------------

Private Sub Dashboard_RenderTopDelays(ByVal ws As Worksheet, ByVal tblWBS As ListObject, ByVal mapWBS As Object, ByVal x As Double, ByVal y As Double, ByVal w As Double, ByVal h As Double)

    Const TOP_DELAY_COUNT As Long = 4

    Dim topIds(1 To 5) As String
    Dim topWbs(1 To 5) As String
    Dim topNames(1 To 5) As String
    Dim topVals(1 To 5) As Double
    Dim arr As Variant
    Dim r As Long
    Dim v As Variant
    Dim i As Long
    Dim lineY As Double
    Dim accentColor As Long

    On Error GoTo RenderFailed

    If tblWBS Is Nothing Then GoTo NoData
    If mapWBS Is Nothing Then GoTo NoData
    If tblWBS.DataBodyRange Is Nothing Then GoTo NoData
    If Not mapWBS.Exists(VTS_COL_FINISH_VARIANCE) Then GoTo NoData

    arr = tblWBS.DataBodyRange.value
    For r = 1 To UBound(arr, 1)
        v = Dashboard_ArrayVal(arr, mapWBS, r, "Finish Variance")
        If IsNumeric(v) Then
            If CDbl(v) > 0 Then
                Dashboard_InsertTopDelay topIds, topWbs, topNames, topVals, _
                    CStr(Dashboard_ArrayVal(arr, mapWBS, r, "ID")), _
                    CStr(Dashboard_ArrayVal(arr, mapWBS, r, "WBS")), _
                    CStr(Dashboard_ArrayVal(arr, mapWBS, r, "Task Name")), _
                    CDbl(v)
            End If
        End If
    Next r

    If topIds(1) <> "" Then
        accentColor = RGB(192, 80, 77)
    Else
        accentColor = RGB(0, 145, 112)
    End If

    Dashboard_AddHotSpotCard ws, x, y, w, h, Dashboard_L("Dérives majeures", "Top Delays"), Dashboard_L("Où est le retard ?", "Where is the delay?"), accentColor

    For i = 1 To TOP_DELAY_COUNT
        lineY = y + 66 + ((i - 1) * 68)
        If topIds(i) <> "" Then
            Dashboard_AddHotSpotRankingItem ws, x + 18, lineY, w - 36, 54, _
                CStr(i), _
                topWbs(i), _
                Dashboard_TruncateText(topNames(i), 30), _
                "+" & CStr(CLng(topVals(i))) & Dashboard_DurationSuffix(), _
                RGB(192, 80, 77)
        End If
    Next i

    If topIds(1) = "" Then
        Dashboard_AddHotSpotText ws, Dashboard_L("Aucune dérive planning détectée", "No schedule delays detected"), x + 18, y + 90, w - 36, 18, RGB(96, 111, 128), 8, False, xlHAlignLeft
    End If

    Exit Sub

NoData:
    Dashboard_AddHotSpotCard ws, x, y, w, h, Dashboard_L("Dérives majeures", "Top Delays"), Dashboard_L("Où est le retard ?", "Where is the delay?"), RGB(160, 170, 181)
    Dashboard_AddHotSpotText ws, Dashboard_L("Aucune donnée de dérive disponible", "No delay data available"), x + 18, y + 90, w - 36, 18, RGB(96, 111, 128), 8, False, xlHAlignLeft
    Exit Sub

RenderFailed:
    Dashboard_WriteHotSpotRuntimeError ws, x + 18, y + 90, w - 36, "Top Delays"

End Sub

'------------------------------------------------------------------------------
' FR: Actualise Render Deadline Health sans modifier les regles metier qui produisent les donnees.
' EN: Refreshes Render Deadline Health without changing the business rules that produce the data.
'------------------------------------------------------------------------------

Private Sub Dashboard_RenderDeadlineHealth(ByVal ws As Worksheet, ByVal tblCalc As ListObject, ByVal mapCalc As Object, ByVal x As Double, ByVal y As Double, ByVal w As Double, ByVal h As Double)

    Dim arr As Variant
    Dim r As Long
    Dim deadlineFloat As Variant
    Dim floatVal As Double
    Dim overdueCount As Long
    Dim onTrackCount As Long
    Dim activeDeadlineCount As Long
    Dim worstFound As Boolean
    Dim riskFound As Boolean
    Dim worstFloat As Double
    Dim riskFloat As Double
    Dim worstWbs As String
    Dim worstName As String
    Dim riskWbs As String
    Dim riskName As String
    Dim accentColor As Long
    Dim riskColor As Long
    Dim riskFill As Long
    Dim closestY As Double

    On Error GoTo RenderFailed

    If tblCalc Is Nothing Then GoTo NoData
    If mapCalc Is Nothing Then GoTo NoData
    If tblCalc.DataBodyRange Is Nothing Then GoTo NoData
    If Not mapCalc.Exists("Deadline") Then GoTo NoData
    If Not mapCalc.Exists("Deadline Float") Then GoTo NoData

    arr = tblCalc.DataBodyRange.value
    For r = 1 To UBound(arr, 1)
        If Dashboard_IsActiveDashboardTask(arr, mapCalc, r) Then
            If HasValue(Dashboard_ArrayVal(arr, mapCalc, r, "Deadline")) Then
                deadlineFloat = Dashboard_ArrayVal(arr, mapCalc, r, "Deadline Float")
                If IsNumeric(deadlineFloat) Then
                    activeDeadlineCount = activeDeadlineCount + 1
                    floatVal = CDbl(deadlineFloat)
                    If floatVal < 0# Then
                        overdueCount = overdueCount + 1
                        If Not worstFound Or floatVal < worstFloat Then
                            worstFound = True
                            worstFloat = floatVal
                            worstWbs = CStr(Dashboard_ArrayVal(arr, mapCalc, r, "WBS"))
                            worstName = CStr(Dashboard_ArrayVal(arr, mapCalc, r, "Task Name"))
                        End If
                    Else
                        onTrackCount = onTrackCount + 1
                        If Not riskFound Or floatVal < riskFloat Then
                            riskFound = True
                            riskFloat = floatVal
                            riskWbs = CStr(Dashboard_ArrayVal(arr, mapCalc, r, "WBS"))
                            riskName = CStr(Dashboard_ArrayVal(arr, mapCalc, r, "Task Name"))
                        End If
                    End If
                End If
            End If
        End If
    Next r

    If activeDeadlineCount = 0 Then
        accentColor = RGB(160, 170, 181)
    ElseIf overdueCount > 0 Then
        accentColor = RGB(192, 80, 77)
    Else
        accentColor = RGB(0, 145, 112)
    End If

    Dashboard_AddHotSpotCard ws, x, y, w, h, Dashboard_L("Santé deadlines", "Deadline Health"), Dashboard_L("Mes engagements sont-ils tenus ?", "Are commitments safe?"), accentColor
    Dashboard_AddHotSpotKpiBlock ws, x + 18, y + 62, (w - 48) / 2, 64, Dashboard_L("En retard", "Overdue"), CStr(overdueCount), IIf(overdueCount > 0, RGB(192, 80, 77), RGB(0, 145, 112)), IIf(overdueCount > 0, RGB(252, 235, 232), RGB(229, 246, 239))
    Dashboard_AddHotSpotKpiBlock ws, x + 30 + ((w - 48) / 2), y + 62, (w - 48) / 2, 64, Dashboard_L("OK", "On Track"), CStr(onTrackCount), RGB(0, 145, 112), RGB(229, 246, 239)

    closestY = y + 152
    If worstFound Then
        Dashboard_AddHotSpotInsightBlock ws, x + 18, closestY, w - 36, 72, Dashboard_L("Plus critique", "Worst Offender"), worstWbs, Dashboard_TruncateText(worstName, 30), CStr(CLng(worstFloat)) & Dashboard_DurationSuffix(), RGB(192, 80, 77), RGB(252, 235, 232)
        closestY = y + 242
    End If

    If riskFound Then
        If riskFloat <= 5# Then
            riskColor = RGB(238, 156, 68)
            riskFill = RGB(255, 244, 226)
        Else
            riskColor = RGB(0, 145, 112)
            riskFill = RGB(229, 246, 239)
        End If
        Dashboard_AddHotSpotInsightBlock ws, x + 18, closestY, w - 36, 72, Dashboard_L("Risque proche", "Closest Risk"), riskWbs, Dashboard_TruncateText(riskName, 30), "+" & CStr(CLng(riskFloat)) & Dashboard_DurationSuffix(), riskColor, riskFill
    ElseIf activeDeadlineCount = 0 Then
        Dashboard_AddHotSpotText ws, Dashboard_L("Aucune deadline active", "No active deadline"), x + 18, closestY + 10, w - 36, 18, RGB(96, 111, 128), 8, False, xlHAlignLeft
    End If

    Exit Sub

NoData:
    Dashboard_AddHotSpotCard ws, x, y, w, h, Dashboard_L("Santé deadlines", "Deadline Health"), Dashboard_L("Mes engagements sont-ils tenus ?", "Are commitments safe?"), RGB(160, 170, 181)
    Dashboard_AddHotSpotText ws, Dashboard_L("Aucune deadline disponible", "No deadline data available"), x + 18, y + 90, w - 36, 18, RGB(96, 111, 128), 8, False, xlHAlignLeft
    Exit Sub

RenderFailed:
    Dashboard_WriteHotSpotRuntimeError ws, x + 18, y + 90, w - 36, "Deadline Health"

End Sub

'------------------------------------------------------------------------------
' FR: Actualise Render Next Milestone sans modifier les regles metier qui produisent les donnees.
' EN: Refreshes Render Next Milestone without changing the business rules that produce the data.
'------------------------------------------------------------------------------

Private Sub Dashboard_RenderNextMilestone(ByVal ws As Worksheet, ByVal tblCalc As ListObject, ByVal mapCalc As Object, ByVal x As Double, ByVal y As Double, ByVal w As Double, ByVal h As Double)

    Dim arr As Variant
    Dim r As Long
    Dim finishVal As Variant
    Dim finishSerial As Double
    Dim todaySerial As Double
    Dim bestFutureFound As Boolean
    Dim bestPastFound As Boolean
    Dim bestFutureDate As Double
    Dim bestPastDate As Double
    Dim bestWbs As String
    Dim bestName As String
    Dim bestDate As Double
    Dim pastWbs As String
    Dim pastName As String
    Dim accentColor As Long
    Dim dateColor As Long

    On Error GoTo RenderFailed

    If tblCalc Is Nothing Then GoTo NoData
    If mapCalc Is Nothing Then GoTo NoData
    If tblCalc.DataBodyRange Is Nothing Then GoTo NoData
    If Not mapCalc.Exists("Task Type") Then GoTo NoData
    If Not mapCalc.Exists("Calculated Finish") Then GoTo NoData

    arr = tblCalc.DataBodyRange.value
    todaySerial = CDbl(Date)

    For r = 1 To UBound(arr, 1)
        If Dashboard_IsMilestoneRow(arr, mapCalc, r) And Dashboard_IsActiveDashboardTask(arr, mapCalc, r) Then
            finishVal = GetCellValue(Dashboard_ArrayVal(arr, mapCalc, r, "Calculated Finish"))
            If Dashboard_HasDateValue(finishVal) Then
                finishSerial = Dashboard_DateNumber(finishVal)
                If finishSerial >= todaySerial Then
                    If Not bestFutureFound Or finishSerial < bestFutureDate Then
                        bestFutureFound = True
                        bestFutureDate = finishSerial
                        bestWbs = CStr(Dashboard_ArrayVal(arr, mapCalc, r, "WBS"))
                        bestName = CStr(Dashboard_ArrayVal(arr, mapCalc, r, "Task Name"))
                    End If
                ElseIf Not bestPastFound Or finishSerial > bestPastDate Then
                    bestPastFound = True
                    bestPastDate = finishSerial
                    pastWbs = CStr(Dashboard_ArrayVal(arr, mapCalc, r, "WBS"))
                    pastName = CStr(Dashboard_ArrayVal(arr, mapCalc, r, "Task Name"))
                End If
            End If
        End If
    Next r

    If bestFutureFound Then
        bestDate = bestFutureDate
        accentColor = RGB(68, 114, 196)
        dateColor = RGB(68, 114, 196)
    ElseIf bestPastFound Then
        bestDate = bestPastDate
        bestWbs = pastWbs
        bestName = pastName
        accentColor = RGB(192, 80, 77)
        dateColor = RGB(192, 80, 77)
    Else
        accentColor = RGB(160, 170, 181)
        dateColor = RGB(96, 111, 128)
    End If

    Dashboard_AddHotSpotCard ws, x, y, w, h, Dashboard_L("Prochain jalon", "Next Milestone"), "", accentColor

    If bestFutureFound Or bestPastFound Then
        Dashboard_AddHotSpotHero ws, Dashboard_TruncateText(bestName, 24), x + 16, y + 58, w - 32, 28, RGB(20, 34, 51)
        Dashboard_AddHotSpotText ws, bestWbs, x + 16, y + 86, w - 32, 14, RGB(96, 111, 128), 8, False, xlHAlignLeft
        Dashboard_AddHotSpotText ws, Dashboard_FormatDate(bestDate, True), x + 16, y + 112, w - 32, 16, RGB(42, 52, 65), 9, False, xlHAlignLeft
        Dashboard_AddHotSpotText ws, Dashboard_DaysRemainingText(CLng(bestDate - todaySerial)), x + 16, y + 136, w - 32, 16, dateColor, 9, True, xlHAlignLeft
    Else
        Dashboard_AddHotSpotText ws, Dashboard_L("Aucun jalon actif", "No active milestone"), x + 16, y + 76, w - 32, 18, RGB(96, 111, 128), 8, False, xlHAlignLeft
    End If

    Exit Sub

NoData:
    Dashboard_AddHotSpotCard ws, x, y, w, h, Dashboard_L("Prochain jalon", "Next Milestone"), "", RGB(160, 170, 181)
    Dashboard_AddHotSpotText ws, Dashboard_L("Aucun jalon disponible", "No milestone data available"), x + 16, y + 76, w - 32, 18, RGB(96, 111, 128), 8, False, xlHAlignLeft
    Exit Sub

RenderFailed:
    Dashboard_WriteHotSpotRuntimeError ws, x + 16, y + 76, w - 32, "Next Milestone"

End Sub

'------------------------------------------------------------------------------
' FR: Actualise Render Next Critical Activity sans modifier les regles metier qui produisent les donnees.
' EN: Refreshes Render Next Critical Activity without changing the business rules that produce the data.
'------------------------------------------------------------------------------

Private Sub Dashboard_RenderNextCriticalActivity(ByVal ws As Worksheet, ByVal tblCalc As ListObject, ByVal mapCalc As Object, ByVal x As Double, ByVal y As Double, ByVal w As Double, ByVal h As Double)

    Dim arr As Variant
    Dim r As Long
    Dim startVal As Variant
    Dim startSerial As Double
    Dim todaySerial As Double
    Dim bestFound As Boolean
    Dim bestDate As Double
    Dim bestWbs As String
    Dim bestName As String
    Dim bestFloat As Variant
    Dim fallbackFound As Boolean
    Dim fallbackDate As Double
    Dim fallbackWbs As String
    Dim fallbackName As String
    Dim fallbackFloat As Variant
    Dim accentColor As Long
    Dim dateColor As Long
    Dim daysToStart As Double

    On Error GoTo RenderFailed

    If tblCalc Is Nothing Then GoTo NoData
    If mapCalc Is Nothing Then GoTo NoData
    If tblCalc.DataBodyRange Is Nothing Then GoTo NoData
    If Not mapCalc.Exists("Critical Path") Then GoTo NoData
    If Not mapCalc.Exists("Calculated Start") Then GoTo NoData

    arr = tblCalc.DataBodyRange.value
    todaySerial = CDbl(Date)

    For r = 1 To UBound(arr, 1)
        If Dashboard_IsActiveDashboardTask(arr, mapCalc, r) Then
            If UCase$(Trim$(CStr(Dashboard_ArrayVal(arr, mapCalc, r, "Critical Path")))) = "CRITICAL" Then
                If Not Dashboard_IsMilestoneRow(arr, mapCalc, r) Then
                    startVal = GetCellValue(Dashboard_ArrayVal(arr, mapCalc, r, "Calculated Start"))
                    If Dashboard_HasDateValue(startVal) Then
                        startSerial = Dashboard_DateNumber(startVal)
                        If startSerial >= todaySerial Then
                            If Not bestFound Or startSerial < bestDate Then
                                bestFound = True
                                bestDate = startSerial
                                bestWbs = CStr(Dashboard_ArrayVal(arr, mapCalc, r, "WBS"))
                                bestName = CStr(Dashboard_ArrayVal(arr, mapCalc, r, "Task Name"))
                                If mapCalc.Exists("Total Float") Then bestFloat = Dashboard_ArrayVal(arr, mapCalc, r, "Total Float") Else bestFloat = Empty
                            End If
                        ElseIf Not fallbackFound Or startSerial > fallbackDate Then
                            fallbackFound = True
                            fallbackDate = startSerial
                            fallbackWbs = CStr(Dashboard_ArrayVal(arr, mapCalc, r, "WBS"))
                            fallbackName = CStr(Dashboard_ArrayVal(arr, mapCalc, r, "Task Name"))
                            If mapCalc.Exists("Total Float") Then fallbackFloat = Dashboard_ArrayVal(arr, mapCalc, r, "Total Float") Else fallbackFloat = Empty
                        End If
                    End If
                End If
            End If
        End If
    Next r

    If Not bestFound And fallbackFound Then
        bestFound = True
        bestDate = fallbackDate
        bestWbs = fallbackWbs
        bestName = fallbackName
        bestFloat = fallbackFloat
    End If

    accentColor = RGB(160, 170, 181)
    dateColor = RGB(96, 111, 128)
    If bestFound Then
        daysToStart = bestDate - todaySerial
        If (IsNumeric(bestFloat) And CDbl(bestFloat) <= 0#) Or daysToStart <= 0# Then
            accentColor = RGB(192, 80, 77)
            dateColor = RGB(192, 80, 77)
        ElseIf daysToStart <= 14# Then
            accentColor = RGB(238, 156, 68)
            dateColor = RGB(238, 156, 68)
        Else
            accentColor = RGB(68, 114, 196)
            dateColor = RGB(68, 114, 196)
        End If
    End If

    Dashboard_AddHotSpotCard ws, x, y, w, h, Dashboard_L("Activité critique", "Next Critical Activity"), "", accentColor

    If bestFound Then
        Dashboard_AddHotSpotHero ws, Dashboard_TruncateText(bestName, 24), x + 16, y + 58, w - 32, 28, RGB(20, 34, 51)
        Dashboard_AddHotSpotText ws, bestWbs, x + 16, y + 86, w - 32, 14, RGB(96, 111, 128), 8, False, xlHAlignLeft
        Dashboard_AddHotSpotText ws, Dashboard_L("Début ", "Starts ") & Dashboard_FormatDate(bestDate, True), x + 16, y + 112, w - 32, 16, RGB(42, 52, 65), 9, False, xlHAlignLeft
        Dashboard_AddHotSpotText ws, Dashboard_DaysRemainingText(CLng(bestDate - todaySerial)), x + 16, y + 136, w - 32, 16, dateColor, 9, True, xlHAlignLeft
        If IsNumeric(bestFloat) Then Dashboard_AddHotSpotText ws, Dashboard_L("Marge: ", "Float: ") & CStr(CLng(bestFloat)) & Dashboard_DurationSuffix(), x + 16, y + h - 24, w - 32, 16, RGB(96, 111, 128), 8, False, xlHAlignLeft
    Else
        Dashboard_AddHotSpotText ws, Dashboard_L("Aucune activité critique active", "No active critical activity"), x + 16, y + 76, w - 32, 18, RGB(96, 111, 128), 8, False, xlHAlignLeft
    End If

    Exit Sub

NoData:
    Dashboard_AddHotSpotCard ws, x, y, w, h, Dashboard_L("Activité critique", "Next Critical Activity"), "", RGB(160, 170, 181)
    Dashboard_AddHotSpotText ws, Dashboard_L("Aucune activité critique disponible", "No critical activity data available"), x + 16, y + 76, w - 32, 18, RGB(96, 111, 128), 8, False, xlHAlignLeft
    Exit Sub

RenderFailed:
    Dashboard_WriteHotSpotRuntimeError ws, x + 16, y + 76, w - 32, "Next Critical Activity"

End Sub

'------------------------------------------------------------------------------
' FR: Emet un checkpoint de diagnostic pour la construction des hot spots Dashboard lorsque la trace est active.
' EN: Emits a diagnostic checkpoint for Dashboard hot-spot construction when tracing is enabled.
'------------------------------------------------------------------------------

Private Sub Dashboard_HotSpotTrace(ByVal helperName As String, ByVal stepName As String)

    gHotSpotHelper = helperName
    gHotSpotStep = stepName
    Debug.Print "Dashboard Hot Spots | " & helperName & " | " & stepName

End Sub

'------------------------------------------------------------------------------
' FR: Ecrit ou synchronise Write Hot Spot Runtime Error dans le stockage possede par le domaine.
' EN: Writes or synchronizes Write Hot Spot Runtime Error in the store owned by the domain.
'------------------------------------------------------------------------------

Private Sub Dashboard_WriteHotSpotRuntimeError( _
    ByVal ws As Worksheet, _
    ByVal x As Double, _
    ByVal y As Double, _
    ByVal w As Double, _
    ByVal cardName As String)

    Dim errNum As Long
    Dim errDesc As String
    Dim msg As String

    errNum = Err.Number
    errDesc = Err.Description
    msg = Dashboard_L("Erreur rendu ", "Render error ") & CStr(errNum) & " - " & errDesc & " | " & cardName & " | " & gHotSpotHelper & " | " & gHotSpotStep
    Debug.Print "Dashboard Hot Spots | " & cardName & " | " & gHotSpotHelper & " | " & gHotSpotStep & " | RenderFailed " & CStr(errNum) & ": " & errDesc

    On Error Resume Next
    Dashboard_AddHotSpotText ws, msg, x, y, w, 28, RGB(192, 80, 77), 8, False, xlHAlignLeft

End Sub

'------------------------------------------------------------------------------
' FR: Ajoute la reference Hot Spot Card a la structure cible fournie par l'appelant.
' EN: Adds the Hot Spot Card reference to the target structure supplied by the caller.
' FR - Effet de bord : cree ou met a jour des shapes Excel.
' EN - Side effect: creates or updates Excel shapes.
'------------------------------------------------------------------------------

Private Sub Dashboard_AddHotSpotCard( _
    ByVal ws As Worksheet, _
    ByVal x As Double, _
    ByVal y As Double, _
    ByVal w As Double, _
    ByVal h As Double, _
    ByVal titleText As String, _
    ByVal subtitleText As String, _
    ByVal accentColor As Long)

    Dim bg As Shape
    Dim accent As Shape

    Dashboard_HotSpotTrace "Dashboard_AddHotSpotCard", "Add background shape"
    Set bg = ws.Shapes.AddShape(msoShapeRoundedRectangle, x, y, w, h)

    Dashboard_HotSpotTrace "Dashboard_AddHotSpotCard", "Format background"
    bg.Name = DASH_PREFIX & "HOTSPOT_CARD"
    bg.Fill.ForeColor.RGB = RGB(255, 255, 255)
    bg.Line.ForeColor.RGB = RGB(224, 230, 237)

    Dashboard_HotSpotTrace "Dashboard_AddHotSpotCard", "Add accent shape"
    Set accent = ws.Shapes.AddShape(msoShapeRectangle, x, y, 5, h)

    Dashboard_HotSpotTrace "Dashboard_AddHotSpotCard", "Format accent"
    accent.Name = DASH_PREFIX & "HOTSPOT_ACCENT"
    accent.Fill.ForeColor.RGB = accentColor
    accent.Line.Visible = msoFalse

    Dashboard_HotSpotTrace "Dashboard_AddHotSpotCard", "Add title text"
    Dashboard_AddHotSpotText ws, titleText, x + 16, y + 11, w - 30, 18, RGB(20, 34, 51), 11, True, xlHAlignLeft

    If Len(Trim$(subtitleText)) > 0 Then
        Dashboard_HotSpotTrace "Dashboard_AddHotSpotCard", "Add subtitle text"
        Dashboard_AddHotSpotText ws, subtitleText, x + 16, y + 32, w - 30, 14, RGB(126, 139, 153), 7, False, xlHAlignLeft
    End If

End Sub

'------------------------------------------------------------------------------
' FR: Ajoute la reference Empty Hot Spot Card a la structure cible fournie par l'appelant.
' EN: Adds the Empty Hot Spot Card reference to the target structure supplied by the caller.
'------------------------------------------------------------------------------

Private Sub Dashboard_AddEmptyHotSpotCard(ByVal ws As Worksheet, ByVal x As Double, ByVal y As Double, ByVal w As Double, ByVal h As Double, ByVal titleText As String)

    Dashboard_AddHotSpotCard ws, x, y, w, h, titleText, "", RGB(160, 170, 181)
    Dashboard_AddHotSpotText ws, Dashboard_NoProjectLoadedText(), x + 16, y + 58, w - 32, 18, RGB(96, 111, 128), 9, True, xlHAlignLeft

End Sub

'------------------------------------------------------------------------------
' FR: Ajoute la reference Hot Spot Kpi Block a la structure cible fournie par l'appelant.
' EN: Adds the Hot Spot Kpi Block reference to the target structure supplied by the caller.
'------------------------------------------------------------------------------

Private Sub Dashboard_AddHotSpotKpiBlock( _
    ByVal ws As Worksheet, _
    ByVal x As Double, _
    ByVal y As Double, _
    ByVal w As Double, _
    ByVal h As Double, _
    ByVal labelText As String, _
    ByVal valueText As String, _
    ByVal colorVal As Long, _
    ByVal fillColor As Long)

    Dashboard_HotSpotTrace "Dashboard_AddHotSpotKpiBlock", "Add KPI background"
    Dashboard_AddHotSpotBlock ws, x, y, w, h, fillColor, colorVal
    Dashboard_AddHotSpotText ws, valueText, x + 12, y + 10, w - 24, 28, colorVal, 18, True, xlHAlignLeft
    Dashboard_AddHotSpotText ws, labelText, x + 12, y + 40, w - 24, 14, RGB(96, 111, 128), 8, False, xlHAlignLeft

End Sub

'------------------------------------------------------------------------------
' FR: Ajoute la reference Hot Spot Insight Block a la structure cible fournie par l'appelant.
' EN: Adds the Hot Spot Insight Block reference to the target structure supplied by the caller.
'------------------------------------------------------------------------------

Private Sub Dashboard_AddHotSpotInsightBlock( _
    ByVal ws As Worksheet, _
    ByVal x As Double, _
    ByVal y As Double, _
    ByVal w As Double, _
    ByVal h As Double, _
    ByVal labelText As String, _
    ByVal wbsText As String, _
    ByVal nameText As String, _
    ByVal metricText As String, _
    ByVal colorVal As Long, _
    ByVal fillColor As Long)

    Dashboard_HotSpotTrace "Dashboard_AddHotSpotInsightBlock", "Add insight background"
    Dashboard_AddHotSpotBlock ws, x, y, w, h, fillColor, colorVal
    Dashboard_AddHotSpotText ws, labelText, x + 12, y + 8, w - 24, 13, RGB(96, 111, 128), 7, True, xlHAlignLeft
    If Len(wbsText) > 0 Then Dashboard_AddHotSpotText ws, wbsText, x + 12, y + 27, 72, 14, colorVal, 8, True, xlHAlignLeft
    Dashboard_AddHotSpotText ws, nameText, x + 12, y + 43, w - 86, 16, RGB(20, 34, 51), 8, False, xlHAlignLeft
    If Len(metricText) > 0 Then Dashboard_AddHotSpotText ws, metricText, x + w - 70, y + 30, 58, 18, colorVal, 10, True, xlHAlignRight

End Sub

'------------------------------------------------------------------------------
' FR: Ajoute la reference Hot Spot Ranking Item a la structure cible fournie par l'appelant.
' EN: Adds the Hot Spot Ranking Item reference to the target structure supplied by the caller.
'------------------------------------------------------------------------------

Private Sub Dashboard_AddHotSpotRankingItem( _
    ByVal ws As Worksheet, _
    ByVal x As Double, _
    ByVal y As Double, _
    ByVal w As Double, _
    ByVal h As Double, _
    ByVal rankText As String, _
    ByVal wbsText As String, _
    ByVal nameText As String, _
    ByVal metricText As String, _
    ByVal colorVal As Long)

    Dashboard_HotSpotTrace "Dashboard_AddHotSpotRankingItem", "Add ranking background"
    Dashboard_AddHotSpotBlock ws, x, y, w, h, RGB(249, 251, 253), colorVal
    Dashboard_AddHotSpotPill ws, rankText, x + 12, y + 13, 28, 28, Dashboard_LightenColor(colorVal, 0.86), colorVal, 9, True
    Dashboard_AddHotSpotText ws, wbsText, x + 50, y + 9, w - 124, 13, RGB(96, 111, 128), 7, False, xlHAlignLeft
    Dashboard_AddHotSpotText ws, nameText, x + 50, y + 26, w - 124, 16, RGB(20, 34, 51), 9, False, xlHAlignLeft
    Dashboard_AddHotSpotPill ws, metricText, x + w - 72, y + 15, 58, 24, Dashboard_LightenColor(colorVal, 0.88), colorVal, 8, True

End Sub

'------------------------------------------------------------------------------
' FR: Ajoute la reference Hot Spot Metric a la structure cible fournie par l'appelant.
' EN: Adds the Hot Spot Metric reference to the target structure supplied by the caller.
'------------------------------------------------------------------------------

Private Sub Dashboard_AddHotSpotMetric( _
    ByVal ws As Worksheet, _
    ByVal x As Double, _
    ByVal y As Double, _
    ByVal w As Double, _
    ByVal labelText As String, _
    ByVal valueText As String, _
    ByVal colorVal As Long)

    Dashboard_AddHotSpotKpiBlock ws, x, y, w, 62, labelText, valueText, colorVal, Dashboard_LightenColor(colorVal, 0.9)

End Sub

'------------------------------------------------------------------------------
' FR: Ajoute la reference Hot Spot Label a la structure cible fournie par l'appelant.
' EN: Adds the Hot Spot Label reference to the target structure supplied by the caller.
'------------------------------------------------------------------------------

Private Sub Dashboard_AddHotSpotLabel(ByVal ws As Worksheet, ByVal txt As String, ByVal x As Double, ByVal y As Double, ByVal w As Double)

    Dashboard_HotSpotTrace "Dashboard_AddHotSpotLabel", "Add label text"
    Dashboard_AddHotSpotText ws, txt, x, y, w, 14, RGB(96, 111, 128), 8, True, xlHAlignLeft

End Sub

'------------------------------------------------------------------------------
' FR: Ajoute la reference Hot Spot Hero a la structure cible fournie par l'appelant.
' EN: Adds the Hot Spot Hero reference to the target structure supplied by the caller.
'------------------------------------------------------------------------------

Private Sub Dashboard_AddHotSpotHero( _
    ByVal ws As Worksheet, _
    ByVal txt As String, _
    ByVal x As Double, _
    ByVal y As Double, _
    ByVal w As Double, _
    ByVal h As Double, _
    ByVal colorVal As Long)

    Dashboard_HotSpotTrace "Dashboard_AddHotSpotHero", "Add hero text"
    Dashboard_AddHotSpotText ws, txt, x, y, w, h, colorVal, 13, True, xlHAlignLeft

End Sub

'------------------------------------------------------------------------------
' FR: Ajoute la collection Hot Spot List Item a la structure cible fournie par l'appelant.
' EN: Adds the Hot Spot List Item collection to the target structure supplied by the caller.
'------------------------------------------------------------------------------

Private Sub Dashboard_AddHotSpotListItem( _
    ByVal ws As Worksheet, _
    ByVal x As Double, _
    ByVal y As Double, _
    ByVal w As Double, _
    ByVal wbsText As String, _
    ByVal nameText As String, _
    ByVal metricText As String, _
    ByVal metricColor As Long)

    Dashboard_AddHotSpotRankingItem ws, x, y, w, 48, "", wbsText, nameText, metricText, metricColor

End Sub

'------------------------------------------------------------------------------
' FR: Ajoute la reference Hot Spot Block a la structure cible fournie par l'appelant.
' EN: Adds the Hot Spot Block reference to the target structure supplied by the caller.
' FR - Effet de bord : cree ou met a jour des shapes Excel.
' EN - Side effect: creates or updates Excel shapes.
'------------------------------------------------------------------------------

Private Sub Dashboard_AddHotSpotBlock( _
    ByVal ws As Worksheet, _
    ByVal x As Double, _
    ByVal y As Double, _
    ByVal w As Double, _
    ByVal h As Double, _
    ByVal fillColor As Long, _
    ByVal lineColor As Long)

    Dim shp As Shape

    Dashboard_HotSpotTrace "Dashboard_AddHotSpotBlock", "Add block shape"
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, x, y, w, h)
    shp.Name = DASH_PREFIX & "HOTSPOT_BLOCK"
    shp.Fill.ForeColor.RGB = fillColor
    shp.Line.ForeColor.RGB = Dashboard_LightenColor(lineColor, 0.7)

End Sub

'------------------------------------------------------------------------------
' FR: Ajoute la reference Hot Spot Pill a la structure cible fournie par l'appelant.
' EN: Adds the Hot Spot Pill reference to the target structure supplied by the caller.
' FR - Effet de bord : cree ou met a jour des shapes Excel.
' EN - Side effect: creates or updates Excel shapes.
'------------------------------------------------------------------------------

Private Sub Dashboard_AddHotSpotPill( _
    ByVal ws As Worksheet, _
    ByVal txt As String, _
    ByVal x As Double, _
    ByVal y As Double, _
    ByVal w As Double, _
    ByVal h As Double, _
    ByVal fillColor As Long, _
    ByVal colorVal As Long, _
    ByVal fontSize As Double, _
    ByVal isBold As Boolean)

    Dim shp As Shape

    Dashboard_HotSpotTrace "Dashboard_AddHotSpotPill", "Add pill shape"
    Set shp = ws.Shapes.AddShape(msoShapeOval, x, y, w, h)
    shp.Name = DASH_PREFIX & "HOTSPOT_PILL"
    shp.Fill.ForeColor.RGB = fillColor
    shp.Line.Visible = msoFalse
    Dashboard_AddHotSpotText ws, txt, x, y + 1, w, h - 2, colorVal, fontSize, isBold, xlHAlignCenter

End Sub

'------------------------------------------------------------------------------
' FR: Ajoute la reference Hot Spot Text a la structure cible fournie par l'appelant.
' EN: Adds the Hot Spot Text reference to the target structure supplied by the caller.
' FR - Effet de bord : cree ou met a jour des shapes Excel.
' EN - Side effect: creates or updates Excel shapes.
'------------------------------------------------------------------------------

Private Sub Dashboard_AddHotSpotText( _
    ByVal ws As Worksheet, _
    ByVal txt As String, _
    ByVal x As Double, _
    ByVal y As Double, _
    ByVal w As Double, _
    ByVal h As Double, _
    ByVal colorVal As Long, _
    ByVal fontSize As Double, _
    ByVal isBold As Boolean, _
    ByVal horizontalAlign As Long)

    Dim shp As Shape

    Dashboard_HotSpotTrace "Dashboard_AddHotSpotText", "Add textbox"
    Set shp = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, x, y, w, h)

    Dashboard_HotSpotTrace "Dashboard_AddHotSpotText", "Format shape fill line"
    shp.Name = DASH_PREFIX & "HOTSPOT_TXT"
    shp.Fill.Visible = msoFalse
    shp.Line.Visible = msoFalse

    Dashboard_HotSpotTrace "Dashboard_AddHotSpotText", "Format legacy TextFrame"
    With shp.TextFrame
        .MarginLeft = 0
        .MarginRight = 0
        .MarginTop = 0
        .MarginBottom = 0
        .HorizontalAlignment = horizontalAlign
        .VerticalAlignment = xlVAlignCenter
        .Characters.Text = txt
        .Characters.Font.Name = "Segoe UI"
        .Characters.Font.Size = fontSize
        .Characters.Font.Color = colorVal
        .Characters.Font.Bold = isBold
    End With

End Sub

'------------------------------------------------------------------------------
' FR: Retourne la valeur Lighten Color sans modifier les donnees d'entree.
' EN: Returns the Lighten Color value without mutating input data.
'------------------------------------------------------------------------------

Private Function Dashboard_LightenColor(ByVal colorVal As Long, ByVal ratio As Double) As Long

    Dim r As Long
    Dim g As Long
    Dim b As Long

    r = colorVal Mod 256
    g = (colorVal \ 256) Mod 256
    b = (colorVal \ 65536) Mod 256

    r = CLng(r + ((255 - r) * ratio))
    g = CLng(g + ((255 - g) * ratio))
    b = CLng(b + ((255 - b) * ratio))

    Dashboard_LightenColor = RGB(r, g, b)

End Function
'------------------------------------------------------------------------------
' FR: Indique si la map Active Dashboard Task satisfait la condition attendue, sans modifier les donnees source.
' EN: Returns whether the Active Dashboard Task map satisfies the expected condition without mutating source data.
'------------------------------------------------------------------------------

Private Function Dashboard_IsActiveDashboardTask(ByRef arr As Variant, ByVal mapTbl As Object, ByVal r As Long) As Boolean

    Dim progressVal As Variant
    Dim actualFinish As Variant

    If mapTbl.Exists("Actual Finish") Then
        actualFinish = GetCellValue(Dashboard_ArrayVal(arr, mapTbl, r, "Actual Finish"))
        If HasValue(actualFinish) Then Exit Function
    End If

    If mapTbl.Exists("% Progress") Then
        progressVal = GetCellValue(Dashboard_ArrayVal(arr, mapTbl, r, "% Progress"))
        If IsNumeric(progressVal) Then
            If Dashboard_NormalizeProgress(progressVal) >= 1# Then Exit Function
        End If
    End If

    Dashboard_IsActiveDashboardTask = True

End Function

'------------------------------------------------------------------------------
' FR: Indique si la map Milestone Row satisfait la condition attendue, sans modifier les donnees source.
' EN: Returns whether the Milestone Row map satisfies the expected condition without mutating source data.
'------------------------------------------------------------------------------

Private Function Dashboard_IsMilestoneRow(ByRef arr As Variant, ByVal mapTbl As Object, ByVal r As Long) As Boolean

    Dim taskTypeVal As String
    Dim durationVal As Variant

    If mapTbl.Exists("Task Type") Then
        taskTypeVal = UCase$(Trim$(CStr(Dashboard_ArrayVal(arr, mapTbl, r, "Task Type"))))
        Dashboard_IsMilestoneRow = (taskTypeVal = "MILESTONE")
        Exit Function
    End If

    If mapTbl.Exists("Calculated Duration") Then
        durationVal = GetCellValue(Dashboard_ArrayVal(arr, mapTbl, r, "Calculated Duration"))
        If IsNumeric(durationVal) Then Dashboard_IsMilestoneRow = (CDbl(durationVal) <= 1#)
    End If

End Function

'------------------------------------------------------------------------------
' FR: Retourne la valeur Days Remaining Text sans modifier les donnees d'entree.
' EN: Returns the Days Remaining Text value without mutating input data.
'------------------------------------------------------------------------------

Private Function Dashboard_DaysRemainingText(ByVal daysVal As Long) As String

    If daysVal < 0 Then
        Dashboard_DaysRemainingText = CStr(Abs(daysVal)) & Dashboard_L("j de retard", " days overdue")
    ElseIf daysVal = 0 Then
        Dashboard_DaysRemainingText = Dashboard_L("aujourd'hui", "today")
    Else
        Dashboard_DaysRemainingText = CStr(daysVal) & Dashboard_L("j restants", " days remaining")
    End If

End Function
'------------------------------------------------------------------------------
' FR: Actualise Render Forecast Issues sans modifier les regles metier qui produisent les donnees.
' EN: Refreshes Render Forecast Issues without changing the business rules that produce the data.
'------------------------------------------------------------------------------

Private Sub Dashboard_RenderForecastIssues(ByVal ws As Worksheet, ByVal tblCalc As ListObject, ByVal mapCalc As Object, ByVal x As Double, ByVal y As Double, ByVal w As Double, ByVal h As Double)

    Dim arr As Variant
    Dim r As Long
    Dim shown As Long
    Dim errMsg As String
    Dim lineY As Double

    Dashboard_AddTableTitle ws, Dashboard_L("Alertes forecast", "Forecast Issues"), x, y, w
    Dashboard_WriteHotspotHeader ws, x, y + 24, Array("WBS", Dashboard_L("Tâche", "Task"), Dashboard_L("Alerte", "Issue"))

    If tblCalc Is Nothing Then Exit Sub
    If tblCalc.DataBodyRange Is Nothing Then Exit Sub
    If Not mapCalc.Exists("ErrorMsg") Then Exit Sub

    arr = tblCalc.DataBodyRange.value
    For r = 1 To UBound(arr, 1)
        errMsg = Trim$(CStr(Dashboard_ArrayVal(arr, mapCalc, r, "ErrorMsg")))
        If InStr(1, errMsg, "Forecast", vbTextCompare) > 0 Then
            shown = shown + 1
            If shown > 5 Then Exit For
            lineY = y + 44 + ((shown - 1) * 24)
            Dashboard_WriteHotspotRow ws, x, lineY, _
                CStr(Dashboard_ArrayVal(arr, mapCalc, r, "WBS")), _
                Left$(CStr(Dashboard_ArrayVal(arr, mapCalc, r, "Task Name")), 18), _
                "Forecast", RGB(192, 80, 77)
        End If
    Next r

    If shown = 0 Then
        Dashboard_WriteEmptyState ws, x, y + 54, w, Dashboard_L("Aucune alerte forecast détectée", "No forecast issues detected")
    End If

End Sub

'------------------------------------------------------------------------------
' FR: Retourne la map Max Date sans modifier les donnees d'entree.
' EN: Returns the Max Date map without mutating input data.
'------------------------------------------------------------------------------

Private Function Dashboard_MaxDate(ByVal tbl As ListObject, ByVal mapTbl As Object, ByVal colName As String) As Variant

    Dim arr As Variant
    Dim r As Long
    Dim v As Variant
    Dim hasDate As Boolean
    Dim maxVal As Double

    If tbl Is Nothing Then Exit Function
    If tbl.DataBodyRange Is Nothing Then Exit Function
    If Not mapTbl.Exists(colName) Then Exit Function

    arr = tbl.DataBodyRange.value
    For r = 1 To UBound(arr, 1)
        v = GetCellValue(arr(r, mapTbl(colName)))
        If Dashboard_HasDateValue(v) Then
            If Not hasDate Then
                maxVal = Dashboard_DateNumber(v)
                hasDate = True
            ElseIf Dashboard_DateNumber(v) > maxVal Then
                maxVal = Dashboard_DateNumber(v)
            End If
        End If
    Next r

    If hasDate Then Dashboard_MaxDate = maxVal

End Function

'------------------------------------------------------------------------------
' FR: Retourne la map Min Date Across sans modifier les donnees d'entree.
' EN: Returns the Min Date Across map without mutating input data.
'------------------------------------------------------------------------------

Private Function Dashboard_MinDateAcross(ByVal tbl As ListObject, ByVal mapTbl As Object, ByVal colNames As Variant) As Variant

    Dim arr As Variant
    Dim r As Long
    Dim c As Variant
    Dim v As Variant
    Dim hasDate As Boolean
    Dim minVal As Double

    If tbl Is Nothing Then Exit Function
    If tbl.DataBodyRange Is Nothing Then Exit Function
    arr = tbl.DataBodyRange.value

    For r = 1 To UBound(arr, 1)
        For Each c In colNames
            If mapTbl.Exists(CStr(c)) Then
                v = GetCellValue(arr(r, mapTbl(CStr(c))))
                If Dashboard_HasDateValue(v) Then
                    If Not hasDate Then
                        minVal = Dashboard_DateNumber(v)
                        hasDate = True
                    ElseIf Dashboard_DateNumber(v) < minVal Then
                        minVal = Dashboard_DateNumber(v)
                    End If
                End If
            End If
        Next c
    Next r

    If hasDate Then Dashboard_MinDateAcross = minVal

End Function

'------------------------------------------------------------------------------
' FR: Retourne la map Max Date Across sans modifier les donnees d'entree.
' EN: Returns the Max Date Across map without mutating input data.
'------------------------------------------------------------------------------

Private Function Dashboard_MaxDateAcross(ByVal tbl As ListObject, ByVal mapTbl As Object, ByVal colNames As Variant) As Variant

    Dim arr As Variant
    Dim r As Long
    Dim c As Variant
    Dim v As Variant
    Dim hasDate As Boolean
    Dim maxVal As Double

    If tbl Is Nothing Then Exit Function
    If tbl.DataBodyRange Is Nothing Then Exit Function
    arr = tbl.DataBodyRange.value

    For r = 1 To UBound(arr, 1)
        For Each c In colNames
            If mapTbl.Exists(CStr(c)) Then
                v = GetCellValue(arr(r, mapTbl(CStr(c))))
                If Dashboard_HasDateValue(v) Then
                    If Not hasDate Then
                        maxVal = Dashboard_DateNumber(v)
                        hasDate = True
                    ElseIf Dashboard_DateNumber(v) > maxVal Then
                        maxVal = Dashboard_DateNumber(v)
                    End If
                End If
            End If
        Next c
    Next r

    If hasDate Then Dashboard_MaxDateAcross = maxVal

End Function

'------------------------------------------------------------------------------
' FR: Retourne la map Numeric Below sans exposer de mutateur sur l'etat source.
' EN: Returns the Numeric Below map without exposing a mutator for source state.
'------------------------------------------------------------------------------

Private Function Dashboard_CountNumericBelow(ByVal tbl As ListObject, ByVal mapTbl As Object, ByVal colName As String, ByVal threshold As Double) As Long

    Dim arr As Variant
    Dim r As Long
    Dim v As Variant

    If tbl Is Nothing Then Exit Function
    If tbl.DataBodyRange Is Nothing Then Exit Function
    If Not mapTbl.Exists(colName) Then Exit Function

    arr = tbl.DataBodyRange.value
    For r = 1 To UBound(arr, 1)
        v = GetCellValue(arr(r, mapTbl(colName)))
        If IsNumeric(v) Then
            If CDbl(v) < threshold Then Dashboard_CountNumericBelow = Dashboard_CountNumericBelow + 1
        End If
    Next r

End Function

'------------------------------------------------------------------------------
' FR: Retourne la map Error Contains sans exposer de mutateur sur l'etat source.
' EN: Returns the Error Contains map without exposing a mutator for source state.
'------------------------------------------------------------------------------

Private Function Dashboard_CountErrorContains(ByVal tbl As ListObject, ByVal mapTbl As Object, ByVal needle As String) As Long

    Dim arr As Variant
    Dim r As Long

    If tbl Is Nothing Then Exit Function
    If tbl.DataBodyRange Is Nothing Then Exit Function
    If Not mapTbl.Exists("ErrorMsg") Then Exit Function

    arr = tbl.DataBodyRange.value
    For r = 1 To UBound(arr, 1)
        If InStr(1, CStr(arr(r, mapTbl("ErrorMsg"))), needle, vbTextCompare) > 0 Then Dashboard_CountErrorContains = Dashboard_CountErrorContains + 1
    Next r

End Function

'------------------------------------------------------------------------------
' FR: Retourne la map Marker sans exposer de mutateur sur l'etat source.
' EN: Returns the Marker map without exposing a mutator for source state.
'------------------------------------------------------------------------------

Private Function Dashboard_CountMarker(ByVal tbl As ListObject, ByVal mapTbl As Object, ByVal colName As String) As Long

    Dim arr As Variant
    Dim r As Long
    Dim txt As String

    If tbl Is Nothing Then Exit Function
    If tbl.DataBodyRange Is Nothing Then Exit Function
    If Not mapTbl.Exists(colName) Then Exit Function

    arr = tbl.DataBodyRange.value
    For r = 1 To UBound(arr, 1)
        txt = UCase$(Trim$(CStr(arr(r, mapTbl(colName)))))
        If txt <> "" And txt <> "NO" And txt <> "FALSE" And txt <> "0" Then Dashboard_CountMarker = Dashboard_CountMarker + 1
    Next r

End Function

'------------------------------------------------------------------------------
' FR: Retourne la map Marker Remaining sans exposer de mutateur sur l'etat source.
' EN: Returns the Marker Remaining map without exposing a mutator for source state.
'------------------------------------------------------------------------------

Private Function Dashboard_CountMarkerRemaining(ByVal tbl As ListObject, ByVal mapTbl As Object, ByVal colName As String) As Long

    Dim arr As Variant
    Dim r As Long
    Dim txt As String
    Dim progressVal As Double
    Dim hasProgress As Boolean
    Dim actualFinish As Variant

    If tbl Is Nothing Then Exit Function
    If tbl.DataBodyRange Is Nothing Then Exit Function
    If Not mapTbl.Exists(colName) Then Exit Function

    hasProgress = mapTbl.Exists("% Progress")

    arr = tbl.DataBodyRange.value
    For r = 1 To UBound(arr, 1)
        txt = UCase$(Trim$(CStr(arr(r, mapTbl(colName)))))
        If txt <> "" And txt <> "NO" And txt <> "FALSE" And txt <> "0" Then
            If hasProgress Then
                progressVal = Dashboard_NormalizeProgress(GetCellValue(arr(r, mapTbl("% Progress"))))
            ElseIf mapTbl.Exists("Actual Finish") Then
                actualFinish = GetCellValue(arr(r, mapTbl("Actual Finish")))
                If Dashboard_HasDateValue(actualFinish) Then progressVal = 1# Else progressVal = 0#
            Else
                progressVal = 0#
            End If
            If progressVal < 1# Then Dashboard_CountMarkerRemaining = Dashboard_CountMarkerRemaining + 1
        End If
    Next r

End Function

'------------------------------------------------------------------------------
' FR: Retourne la valeur Contract Status sans modifier les donnees d'entree.
' EN: Returns the Contract Status value without mutating input data.
'------------------------------------------------------------------------------

Private Function Dashboard_ContractStatus(ByVal driftDays As Variant, ByRef statusColor As Long) As String

    statusColor = RGB(96, 111, 128)

    If Not HasValue(driftDays) Then
        Dashboard_ContractStatus = "UNKNOWN"
    ElseIf CLng(driftDays) <= 0 Then
        Dashboard_ContractStatus = "ON CONTRACT"
        statusColor = RGB(0, 145, 112)
    ElseIf CLng(driftDays) <= 10 Then
        Dashboard_ContractStatus = "MINOR DELAY"
        statusColor = RGB(238, 156, 68)
    Else
        Dashboard_ContractStatus = "CONTRACT DELAY"
        statusColor = RGB(192, 80, 77)
    End If

End Function

'------------------------------------------------------------------------------
' FR: Construit la map Metrics a partir des donnees fournies par l'appelant.
' EN: Builds the Metrics map from data supplied by the caller.
'------------------------------------------------------------------------------

Private Function Dashboard_BuildMetrics( _
    ByVal tblCalc As ListObject, _
    ByVal mapCalc As Object, _
    ByVal scurveProjection As Object) As Object

    Dim metrics As Object
    Dim baselineFinish As Variant
    Dim calcFinish As Variant
    Dim driftDays As Variant
    Dim contractColor As Long

    Set metrics = CreateObject("Scripting.Dictionary")
    metrics("ActualProgress") = CDbl(scurveProjection("ActualProgress"))
    metrics("PlannedProgress") = CDbl(scurveProjection("PlannedProgress"))
    metrics("ProgressVariance") = CDbl(metrics("ActualProgress")) - CDbl(metrics("PlannedProgress"))
    baselineFinish = Dashboard_MaxDate(tblCalc, mapCalc, "Baseline Finish")
    calcFinish = Dashboard_MaxDate(tblCalc, mapCalc, "Calculated Finish")
    metrics("BaselineFinish") = baselineFinish
    metrics("ForecastFinish") = calcFinish
    If Dashboard_HasDateValue(baselineFinish) And Dashboard_HasDateValue(calcFinish) Then
        driftDays = CLng(Dashboard_DateNumber(calcFinish) - Dashboard_DateNumber(baselineFinish))
        metrics("DriftDays") = driftDays
    Else
        metrics("DriftDays") = Empty
    End If
    metrics("DeadlineRiskCount") = Dashboard_CountNumericBelow(tblCalc, mapCalc, "Deadline Float", 0#)
    metrics("ForecastIssueCount") = Dashboard_CountErrorContains(tblCalc, mapCalc, "Forecast")
    metrics("LongestPathCount") = Dashboard_CountMarker(tblCalc, mapCalc, "Longest Path")
    metrics("LongestPathRemaining") = Dashboard_CountMarkerRemaining(tblCalc, mapCalc, "Longest Path")
    metrics("CriticalPathCount") = Dashboard_CountMarker(tblCalc, mapCalc, "Critical Path")
    metrics("ContractStatus") = Dashboard_ContractStatus(metrics("DriftDays"), contractColor)
    Set Dashboard_BuildMetrics = metrics

End Function

'------------------------------------------------------------------------------
' FR: Construit la map Snapshot From Current Data a partir des donnees fournies par l'appelant.
' EN: Builds the Snapshot From Current Data map from data supplied by the caller.
' FR - Effet de bord : ecrit dans une table Excel detenue par le workflow.
' EN - Side effect: writes to an Excel table owned by the workflow.
'------------------------------------------------------------------------------

Private Sub Dashboard_CreateSnapshotFromCurrentData()

    Dim tblCalc As ListObject
    Dim mapCalc As Object
    Dim scurveProjection As Object
    Dim metrics As Object
    Dim tblSnap As ListObject
    Dim newRow As ListRow
    Dim snapshotId As Long
    Dim ts As Date

    Set tblCalc = ThisWorkbook.Worksheets("CALC").ListObjects("tbl_CALC")
    Set mapCalc = CanonicalIdentity_BuildColumnMap(tblCalc)
    Set scurveProjection = SCurve_BuildDashboardProjection()
    Set metrics = Dashboard_BuildMetrics(tblCalc, mapCalc, scurveProjection)
    Set tblSnap = Dashboard_EnsureSnapshotsTable()
    snapshotId = Dashboard_NextSnapshotId(tblSnap)
    ts = Now
    Set newRow = tblSnap.ListRows.Add
    With newRow.Range
        .Cells(1, tblSnap.ListColumns("SnapshotId").Index).value = snapshotId
        .Cells(1, tblSnap.ListColumns("SnapshotDateTime").Index).value = ts
        .Cells(1, tblSnap.ListColumns("SnapshotLabel").Index).value = Dashboard_SnapshotLabel(snapshotId, ts)
        .Cells(1, tblSnap.ListColumns("ActualProgress").Index).value = metrics("ActualProgress")
        .Cells(1, tblSnap.ListColumns("PlannedProgress").Index).value = metrics("PlannedProgress")
        .Cells(1, tblSnap.ListColumns("ProgressVariance").Index).value = metrics("ProgressVariance")
        If Dashboard_HasDateValue(metrics("ForecastFinish")) Then .Cells(1, tblSnap.ListColumns("ForecastFinish").Index).value = CDate(Dashboard_DateNumber(metrics("ForecastFinish")))
        If Dashboard_HasDateValue(metrics("BaselineFinish")) Then .Cells(1, tblSnap.ListColumns("BaselineFinish").Index).value = CDate(Dashboard_DateNumber(metrics("BaselineFinish")))
        If HasValue(metrics("DriftDays")) Then .Cells(1, tblSnap.ListColumns("DriftDays").Index).value = CLng(metrics("DriftDays"))
        .Cells(1, tblSnap.ListColumns("ContractStatus").Index).value = metrics("ContractStatus")
        .Cells(1, tblSnap.ListColumns("DeadlineRiskCount").Index).value = metrics("DeadlineRiskCount")
        .Cells(1, tblSnap.ListColumns("ForecastIssueCount").Index).value = metrics("ForecastIssueCount")
        .Cells(1, tblSnap.ListColumns("LongestPathCount").Index).value = metrics("LongestPathCount")
        .Cells(1, tblSnap.ListColumns("LongestPathRemaining").Index).value = metrics("LongestPathRemaining")
        .Cells(1, tblSnap.ListColumns("CriticalPathCount").Index).value = metrics("CriticalPathCount")
    End With

End Sub

'------------------------------------------------------------------------------
' FR: Cree ou remet en conformite la reference Snapshots Table de maniere idempotente.
' EN: Creates or restores the Snapshots Table reference to a compliant state idempotently.
' FR - Effet de bord : ecrit dans une table Excel detenue par le workflow.
' FR - Effet de bord : efface uniquement les donnees ou objets cibles du contrat.
' EN - Side effect: writes to an Excel table owned by the workflow.
' EN - Side effect: clears only data or objects targeted by the contract.
'------------------------------------------------------------------------------

Private Function Dashboard_EnsureSnapshotsTable() As ListObject

    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim headers As Variant
    Dim i As Long
    Dim nm As Name

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(DASHBOARD_SNAPSHOT_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = DASHBOARD_SNAPSHOT_SHEET
    End If

    On Error Resume Next
    Set tbl = ws.ListObjects(DASHBOARD_SNAPSHOT_TABLE)
    On Error GoTo 0

    If tbl Is Nothing Then
        headers = Array("SnapshotId", "SnapshotDateTime", "SnapshotLabel", "ActualProgress", "PlannedProgress", "ProgressVariance", "ForecastFinish", "BaselineFinish", "DriftDays", "ContractStatus", "DeadlineRiskCount", "ForecastIssueCount", "LongestPathCount", "LongestPathRemaining", "CriticalPathCount")
        For i = LBound(headers) To UBound(headers)
            ws.Cells(1, i + 1).value = CStr(headers(i))
        Next i
        Set tbl = ws.ListObjects.Add(xlSrcRange, ws.Range(ws.Cells(1, 1), ws.Cells(1, UBound(headers) + 1)), , xlYes)
        tbl.Name = DASHBOARD_SNAPSHOT_TABLE
    End If

    On Error Resume Next
    ThisWorkbook.names("DashboardSnapshotLabels").Delete
    On Error GoTo 0
    If Not tbl.DataBodyRange Is Nothing Then
        ThisWorkbook.names.Add Name:="DashboardSnapshotLabels", RefersTo:="=" & DASHBOARD_SNAPSHOT_TABLE & "[SnapshotLabel]"
    End If

    Set Dashboard_EnsureSnapshotsTable = tbl

End Function

'------------------------------------------------------------------------------
' FR: Retourne la reference Next Snapshot ID sans modifier les donnees d'entree.
' EN: Returns the Next Snapshot ID reference without mutating input data.
'------------------------------------------------------------------------------

Private Function Dashboard_NextSnapshotId(ByVal tbl As ListObject) As Long

    Dim arr As Variant
    Dim r As Long
    Dim v As Variant
    Dim maxId As Long

    If tbl Is Nothing Then Exit Function
    If tbl.DataBodyRange Is Nothing Then
        Dashboard_NextSnapshotId = 1
        Exit Function
    End If

    arr = tbl.DataBodyRange.value
    For r = 1 To UBound(arr, 1)
        v = arr(r, tbl.ListColumns("SnapshotId").Index)
        If IsNumeric(v) Then If CLng(v) > maxId Then maxId = CLng(v)
    Next r

    Dashboard_NextSnapshotId = maxId + 1

End Function

'------------------------------------------------------------------------------
' FR: Retourne la valeur Snapshot Label sans modifier les donnees d'entree.
' EN: Returns the Snapshot Label value without mutating input data.
'------------------------------------------------------------------------------

Private Function Dashboard_SnapshotLabel(ByVal snapshotId As Long, ByVal snapshotDate As Date) As String

    Dashboard_SnapshotLabel = "#" & CStr(snapshotId) & " - " & Format$(snapshotDate, "dd/mm/yyyy hh:nn")

End Function

'------------------------------------------------------------------------------
' FR: Retourne la reference Default Snapshot Labels sans exposer de mutateur sur l'etat source.
' EN: Returns the Default Snapshot Labels reference without exposing a mutator for source state.
'------------------------------------------------------------------------------

Private Sub Dashboard_GetDefaultSnapshotLabels(ByVal tbl As ListObject, ByRef fromLabel As String, ByRef toLabel As String)

    Dim n As Long

    fromLabel = ""
    toLabel = ""
    If tbl Is Nothing Then Exit Sub
    If tbl.DataBodyRange Is Nothing Then Exit Sub

    n = tbl.ListRows.Count
    toLabel = CStr(tbl.DataBodyRange.Cells(n, tbl.ListColumns("SnapshotLabel").Index).value)
    If n > 1 Then
        fromLabel = CStr(tbl.DataBodyRange.Cells(n - 1, tbl.ListColumns("SnapshotLabel").Index).value)
    Else
        fromLabel = toLabel
    End If

End Sub

'------------------------------------------------------------------------------
' FR: Retourne la reference Snapshot Label Exists sans modifier les donnees d'entree.
' EN: Returns the Snapshot Label Exists reference without mutating input data.
'------------------------------------------------------------------------------

Private Function Dashboard_SnapshotLabelExists(ByVal tbl As ListObject, ByVal snapshotLabel As String) As Boolean

    Dim arr As Variant
    Dim r As Long

    If tbl Is Nothing Then Exit Function
    If tbl.DataBodyRange Is Nothing Then Exit Function

    arr = tbl.DataBodyRange.value
    For r = 1 To UBound(arr, 1)
        If CStr(arr(r, tbl.ListColumns("SnapshotLabel").Index)) = snapshotLabel Then
            Dashboard_SnapshotLabelExists = True
            Exit Function
        End If
    Next r

End Function

'------------------------------------------------------------------------------
' FR: Retourne la map Momentum Status sans modifier les donnees d'entree.
' EN: Returns the Momentum Status map without mutating input data.
'------------------------------------------------------------------------------

Private Function Dashboard_MomentumStatus(ByVal fromLabel As String, ByVal toLabel As String, ByRef statusColor As Long, ByRef compareText As String, ByRef subText As String) As String

    Dim tbl As ListObject
    Dim fromRow As Object
    Dim toRow As Object
    Dim progressDelta As Double
    Dim forecastDelta As Double
    Dim riskDelta As Double
    Dim fromDefault As String
    Dim toDefault As String

    statusColor = RGB(96, 111, 128)
    compareText = Dashboard_L("Aucun snapshot", "No snapshots")
    subText = Dashboard_L("Historique insuffisant", "Insufficient history")
    Dashboard_MomentumStatus = "INSUFFICIENT HISTORY"

    Set tbl = Dashboard_EnsureSnapshotsTable()
    Dashboard_GetDefaultSnapshotLabels tbl, fromDefault, toDefault
    If fromLabel = "" Then fromLabel = fromDefault
    If toLabel = "" Then toLabel = toDefault

    If fromLabel = "" Or toLabel = "" Then Exit Function

    Set fromRow = Dashboard_GetSnapshotByLabel(tbl, fromLabel)
    Set toRow = Dashboard_GetSnapshotByLabel(tbl, toLabel)
    If fromRow Is Nothing Or toRow Is Nothing Then Exit Function

    compareText = Dashboard_FormatDate(fromRow("SnapshotDateTime"), True) & " " & ChrW(8594) & " " & Dashboard_FormatDate(toRow("SnapshotDateTime"), True)
    If CStr(fromRow("SnapshotId")) = CStr(toRow("SnapshotId")) Then Exit Function

    progressDelta = CDbl(toRow("ActualProgress")) - CDbl(fromRow("ActualProgress"))
    forecastDelta = Dashboard_DateNumber(toRow("ForecastFinish")) - Dashboard_DateNumber(fromRow("ForecastFinish"))
    riskDelta = CDbl(toRow("DeadlineRiskCount")) - CDbl(fromRow("DeadlineRiskCount"))
    subText = Dashboard_L("Avancement ", "Progress ") & Dashboard_FormatPercentSigned(progressDelta) & _
        " | " & Dashboard_L("Fin ", "Forecast ") & Dashboard_FormatSignedCompactDays(forecastDelta) & _
        " | " & Dashboard_L("Risques ", "Risks ") & Dashboard_FormatSignedNumber(riskDelta)

    If forecastDelta > 0 Or riskDelta > 0 Then
        Dashboard_MomentumStatus = "DETERIORATING"
        statusColor = RGB(192, 80, 77)
    ElseIf progressDelta > 0 And (forecastDelta < 0 Or riskDelta < 0) Then
        Dashboard_MomentumStatus = "IMPROVING"
        statusColor = RGB(0, 145, 112)
    Else
        Dashboard_MomentumStatus = "STABLE"
        statusColor = RGB(238, 156, 68)
    End If

End Function

'------------------------------------------------------------------------------
' FR: Retourne la map Snapshot By Label sans exposer de mutateur sur l'etat source.
' EN: Returns the Snapshot By Label map without exposing a mutator for source state.
'------------------------------------------------------------------------------

Private Function Dashboard_GetSnapshotByLabel(ByVal tbl As ListObject, ByVal snapshotLabel As String) As Object

    Dim snap As Object
    Dim arr As Variant
    Dim r As Long
    Dim c As Long

    If tbl Is Nothing Then Exit Function
    If tbl.DataBodyRange Is Nothing Then Exit Function

    arr = tbl.DataBodyRange.value
    For r = 1 To UBound(arr, 1)
        If CStr(arr(r, tbl.ListColumns("SnapshotLabel").Index)) = snapshotLabel Then
            Set snap = CreateObject("Scripting.Dictionary")
            For c = 1 To tbl.ListColumns.Count
                snap(tbl.ListColumns(c).Name) = arr(r, c)
            Next c
            Set Dashboard_GetSnapshotByLabel = snap
            Exit Function
        End If
    Next r

End Function

'------------------------------------------------------------------------------
' FR: Retourne la map Comparison Rows sans exposer de mutateur sur l'etat source.
' EN: Returns the Comparison Rows map without exposing a mutator for source state.
'------------------------------------------------------------------------------

Private Sub Dashboard_GetComparisonRows(ByVal fromLabel As String, ByVal toLabel As String, ByRef fromRow As Object, ByRef toRow As Object)

    Dim tbl As ListObject
    Dim fromDefault As String
    Dim toDefault As String

    Set tbl = Dashboard_EnsureSnapshotsTable()
    Dashboard_GetDefaultSnapshotLabels tbl, fromDefault, toDefault
    If fromLabel = "" Then fromLabel = fromDefault
    If toLabel = "" Then toLabel = toDefault

    Set fromRow = Dashboard_GetSnapshotByLabel(tbl, fromLabel)
    Set toRow = Dashboard_GetSnapshotByLabel(tbl, toLabel)

End Sub

'------------------------------------------------------------------------------
' FR: Retourne la map Comparison Text sans modifier les donnees d'entree.
' EN: Returns the Comparison Text map without mutating input data.
'------------------------------------------------------------------------------

Private Function Dashboard_ComparisonText(ByVal fromRow As Object, ByVal toRow As Object) As String

    If fromRow Is Nothing Or toRow Is Nothing Then
        Dashboard_ComparisonText = Dashboard_L("Comparaison indisponible", "Comparison unavailable")
        Exit Function
    End If

    If Not Dashboard_HasDateValue(fromRow("SnapshotDateTime")) Or Not Dashboard_HasDateValue(toRow("SnapshotDateTime")) Then
        Dashboard_ComparisonText = Dashboard_L("Comparaison indisponible", "Comparison unavailable")
        Exit Function
    End If

    Dashboard_ComparisonText = Dashboard_FormatDate(fromRow("SnapshotDateTime"), True) & _
        " " & ChrW(8594) & " " & _
        Dashboard_FormatDate(toRow("SnapshotDateTime"), True)

End Function

'------------------------------------------------------------------------------
' FR: Retourne la map Progress Delta Text sans modifier les donnees d'entree.
' EN: Returns the Progress Delta Text map without mutating input data.
'------------------------------------------------------------------------------

Private Function Dashboard_ProgressDeltaText(ByVal fromRow As Object, ByVal toRow As Object) As String

    Dim deltaVal As Double

    If fromRow Is Nothing Or toRow Is Nothing Then
        Dashboard_ProgressDeltaText = Dashboard_L("Delta indisponible", "Delta unavailable")
        Exit Function
    End If

    deltaVal = CDbl(toRow("ActualProgress")) - CDbl(fromRow("ActualProgress"))
    Dashboard_ProgressDeltaText = Dashboard_FormatPercentSigned(deltaVal) & " " & Dashboard_L("vs Début", "vs From")

End Function

'------------------------------------------------------------------------------
' FR: Retourne la valeur Behind Plan Text sans modifier les donnees d'entree.
' EN: Returns the Behind Plan Text value without mutating input data.
'------------------------------------------------------------------------------

Private Function Dashboard_BehindPlanText(ByVal plannedProgress As Double, ByVal actualProgress As Double, ByRef statusColor As Long) As String

    Dim behindPlan As Double

    behindPlan = plannedProgress - actualProgress

    If behindPlan <= 0# Then
        statusColor = RGB(0, 145, 112)
        Dashboard_BehindPlanText = Dashboard_L("a l'heure ou en avance", "ahead or on plan")
    ElseIf behindPlan <= 0.1 Then
        statusColor = RGB(238, 156, 68)
        Dashboard_BehindPlanText = Format$(behindPlan, "0%") & Dashboard_L(" de retard vs plan", " behind plan")
    Else
        statusColor = RGB(192, 80, 77)
        Dashboard_BehindPlanText = Format$(behindPlan, "0%") & Dashboard_L(" de retard vs plan", " behind plan")
    End If

End Function

'------------------------------------------------------------------------------
' FR: Retourne la map Forecast Delta Text sans modifier les donnees d'entree.
' EN: Returns the Forecast Delta Text map without mutating input data.
'------------------------------------------------------------------------------

Private Function Dashboard_ForecastDeltaText(ByVal fromRow As Object, ByVal toRow As Object) As String

    Dim deltaDays As Long

    If fromRow Is Nothing Or toRow Is Nothing Then
        Dashboard_ForecastDeltaText = Dashboard_L("Comparaison indisponible", "Comparison unavailable")
        Exit Function
    End If
    If Not Dashboard_HasDateValue(fromRow("ForecastFinish")) Or Not Dashboard_HasDateValue(toRow("ForecastFinish")) Then
        Dashboard_ForecastDeltaText = Dashboard_L("Fin non comparable", "Forecast not comparable")
        Exit Function
    End If

    deltaDays = CLng(Dashboard_DateNumber(toRow("ForecastFinish")) - Dashboard_DateNumber(fromRow("ForecastFinish")))
    If deltaDays < 0 Then
        Dashboard_ForecastDeltaText = Dashboard_L("Fin améliorée de ", "Forecast improved by ") & CStr(Abs(deltaDays)) & Dashboard_L("j", "d")
    ElseIf deltaDays > 0 Then
        Dashboard_ForecastDeltaText = Dashboard_L("Fin décalée de ", "Forecast slipped by ") & CStr(deltaDays) & Dashboard_L("j", "d")
    Else
        Dashboard_ForecastDeltaText = Dashboard_L("Fin stable vs Début", "Forecast stable vs From")
    End If

End Function

'------------------------------------------------------------------------------
' FR: Retourne la valeur Contract Drift Text sans modifier les donnees d'entree.
' EN: Returns the Contract Drift Text value without mutating input data.
'------------------------------------------------------------------------------

Private Function Dashboard_ContractDriftText(ByVal driftDays As Variant, ByRef statusColor As Long) As String

    If Not HasValue(driftDays) Then
        statusColor = RGB(96, 111, 128)
        Dashboard_ContractDriftText = Dashboard_L("dérive contrat indisponible", "contract drift unavailable")
    ElseIf CLng(driftDays) <= 0 Then
        statusColor = RGB(0, 145, 112)
        Dashboard_ContractDriftText = CStr(CLng(driftDays)) & Dashboard_L("j vs contrat", "d vs contract")
    ElseIf CLng(driftDays) <= 30 Then
        statusColor = RGB(238, 156, 68)
        Dashboard_ContractDriftText = "+" & CStr(CLng(driftDays)) & Dashboard_L("j vs contrat", "d vs contract")
    Else
        statusColor = RGB(192, 80, 77)
        Dashboard_ContractDriftText = "+" & CStr(CLng(driftDays)) & Dashboard_L("j vs contrat", "d vs contract")
    End If

End Function

'------------------------------------------------------------------------------
' FR: Retourne la map Critical Activities Hero sans modifier les donnees d'entree.
' EN: Returns the Critical Activities Hero map without mutating input data.
'------------------------------------------------------------------------------

Private Function Dashboard_CriticalActivitiesHero(ByVal fromRow As Object, ByVal toRow As Object, ByRef statusColor As Long, ByRef subText As String) As String

    Dim fromCount As Long
    Dim toCount As Long
    Dim deltaVal As Long

    statusColor = RGB(96, 111, 128)
    subText = Dashboard_L("historique insuffisant", "insufficient history")
    Dashboard_CriticalActivitiesHero = "-"

    If toRow Is Nothing Then Exit Function

    toCount = CLng(toRow("CriticalPathCount"))
    If fromRow Is Nothing Then
        Dashboard_CriticalActivitiesHero = CStr(toCount) & " " & ChrW(8594)
        Exit Function
    End If

    fromCount = CLng(fromRow("CriticalPathCount"))
    deltaVal = toCount - fromCount

    If deltaVal > 0 Then
        statusColor = RGB(192, 80, 77)
        Dashboard_CriticalActivitiesHero = CStr(toCount) & " " & ChrW(8599)
        subText = "+" & CStr(deltaVal) & Dashboard_L(" nouvelles activités critiques", " new critical activities")
    ElseIf deltaVal < 0 Then
        statusColor = RGB(0, 145, 112)
        Dashboard_CriticalActivitiesHero = CStr(toCount) & " " & ChrW(8600)
        subText = CStr(deltaVal) & Dashboard_L(" activités critiques retirées", " critical activities removed")
    Else
        statusColor = RGB(238, 156, 68)
        Dashboard_CriticalActivitiesHero = CStr(toCount) & " " & ChrW(8594)
        subText = Dashboard_L("0 changement activités critiques", "0 critical activities change")
    End If

End Function

'------------------------------------------------------------------------------
' FR: Retourne la map Schedule Momentum Status sans modifier les donnees d'entree.
' EN: Returns the Schedule Momentum Status map without mutating input data.
'------------------------------------------------------------------------------

Private Function Dashboard_ScheduleMomentumStatus(ByVal fromRow As Object, ByVal toRow As Object, ByRef statusColor As Long, ByRef subText As String) As String

    Dim delayFrom As Double
    Dim delayTo As Double
    Dim delayDelta As Double

    statusColor = RGB(96, 111, 128)
    subText = Dashboard_L("Snapshots requis", "Need snapshots")
    Dashboard_ScheduleMomentumStatus = "INSUFFICIENT HISTORY"

    If fromRow Is Nothing Or toRow Is Nothing Then Exit Function
    If CStr(fromRow("SnapshotId")) = CStr(toRow("SnapshotId")) Then Exit Function

    delayFrom = CDbl(fromRow("PlannedProgress")) - CDbl(fromRow("ActualProgress"))
    delayTo = CDbl(toRow("PlannedProgress")) - CDbl(toRow("ActualProgress"))
    delayDelta = delayTo - delayFrom

    If delayDelta < 0 Then
        statusColor = RGB(0, 145, 112)
        Dashboard_ScheduleMomentumStatus = "IMPROVING"
        subText = Dashboard_FormatMomentumDelayPercent(delayDelta) & Dashboard_L(" retard vs snapshot precedent", " delay vs previous snapshot")
    ElseIf delayDelta > 0 Then
        statusColor = RGB(192, 80, 77)
        Dashboard_ScheduleMomentumStatus = "DETERIORATING"
        subText = Dashboard_FormatMomentumDelayPercent(delayDelta) & Dashboard_L(" retard vs snapshot precedent", " delay vs previous snapshot")
    Else
        statusColor = RGB(238, 156, 68)
        Dashboard_ScheduleMomentumStatus = "STABLE"
        subText = Dashboard_L("0% évolution du retard", "0% delay change")
    End If

End Function

'------------------------------------------------------------------------------
' FR: Retourne la valeur Local Contract Status sans modifier les donnees d'entree.
' EN: Returns the Local Contract Status value without mutating input data.
'------------------------------------------------------------------------------

Private Function Dashboard_LocalContractStatus(ByVal statusText As String) As String

    Select Case UCase$(Trim$(statusText))
        Case "ON CONTRACT": Dashboard_LocalContractStatus = Dashboard_L("CONTRAT OK", "ON CONTRACT")
        Case "MINOR DELAY": Dashboard_LocalContractStatus = Dashboard_L("RETARD MINEUR", "MINOR DELAY")
        Case "CONTRACT DELAY": Dashboard_LocalContractStatus = Dashboard_L("RETARD CONTRAT", "CONTRACT DELAY")
        Case Else: Dashboard_LocalContractStatus = statusText
    End Select

End Function

'------------------------------------------------------------------------------
' FR: Retourne la valeur Local Momentum Status sans modifier les donnees d'entree.
' EN: Returns the Local Momentum Status value without mutating input data.
'------------------------------------------------------------------------------

Private Function Dashboard_LocalMomentumStatus(ByVal statusText As String) As String

    Select Case UCase$(Trim$(statusText))
        Case "IMPROVING": Dashboard_LocalMomentumStatus = Dashboard_L("AMÉLIORATION", "IMPROVING")
        Case "DETERIORATING": Dashboard_LocalMomentumStatus = Dashboard_L("DÉGRADATION", "DETERIORATING")
        Case "STABLE": Dashboard_LocalMomentumStatus = Dashboard_L("STABLE", "STABLE")
        Case "INSUFFICIENT HISTORY": Dashboard_LocalMomentumStatus = Dashboard_L("PAS D'HIST.", "NO HISTORY")
        Case Else: Dashboard_LocalMomentumStatus = statusText
    End Select

End Function

'------------------------------------------------------------------------------
' FR: Retourne la map Array Val sans modifier les donnees d'entree.
' EN: Returns the Array Val map without mutating input data.
'------------------------------------------------------------------------------

Private Function Dashboard_ArrayVal(ByRef arr As Variant, ByVal mapTbl As Object, ByVal r As Long, ByVal colName As String) As Variant

    If mapTbl.Exists(colName) Then Dashboard_ArrayVal = arr(r, mapTbl(colName))

End Function

'------------------------------------------------------------------------------
' FR: Indique si la map Summary Array Row satisfait la condition attendue, sans modifier les donnees source.
' EN: Returns whether the Summary Array Row map satisfies the expected condition without mutating source data.
'------------------------------------------------------------------------------

Private Function Dashboard_IsSummaryArrayRow(ByRef arr As Variant, ByVal mapTbl As Object, ByVal r As Long) As Boolean

    Dim v As Variant
    Dim txt As String

    If Not mapTbl.Exists("IsSummary") Then Exit Function
    v = arr(r, mapTbl("IsSummary"))

    If VarType(v) = vbBoolean Then
        Dashboard_IsSummaryArrayRow = CBool(v)
    ElseIf IsNumeric(v) Then
        Dashboard_IsSummaryArrayRow = (CDbl(v) <> 0)
    Else
        txt = UCase$(Trim$(CStr(v)))
        Dashboard_IsSummaryArrayRow = (txt = "TRUE" Or txt = "YES" Or txt = "SUMMARY" Or txt = "1")
    End If

End Function

'------------------------------------------------------------------------------
' FR: Construit l'index Progress By ID a partir des donnees fournies par l'appelant.
' EN: Builds the Progress By ID index from data supplied by the caller.
'------------------------------------------------------------------------------

Private Function Dashboard_BuildProgressById(ByVal tblWBS As ListObject, ByVal mapWBS As Object) As Object

    Dim progressById As Object
    Dim arr As Variant
    Dim r As Long
    Dim idVal As String

    Set progressById = CreateObject("Scripting.Dictionary")

    If tblWBS Is Nothing Then GoTo CleanExit
    If tblWBS.DataBodyRange Is Nothing Then GoTo CleanExit
    If Not mapWBS.Exists(VTS_COL_ID) Then GoTo CleanExit
    If Not mapWBS.Exists(VTS_COL_PROGRESS_PERCENT) Then GoTo CleanExit

    arr = tblWBS.DataBodyRange.value
    For r = 1 To UBound(arr, 1)
        idVal = CStr(GetCellValue(arr(r, mapWBS(VTS_COL_ID))))
        If idVal <> "" Then progressById(idVal) = Dashboard_NormalizeProgress(GetCellValue(arr(r, mapWBS(VTS_COL_PROGRESS_PERCENT))))
    Next r

CleanExit:
    Set Dashboard_BuildProgressById = progressById

End Function

'------------------------------------------------------------------------------
' FR: Retourne la map Summary Progress sans modifier les donnees d'entree.
' EN: Returns the Summary Progress map without mutating input data.
'------------------------------------------------------------------------------

Private Function Dashboard_SummaryProgress(ByRef arr As Variant, ByVal mapTbl As Object, ByVal progressById As Object, ByVal r As Long) As Double

    Dim parentWbs As String
    Dim parentLevel As Long
    Dim childWbs As String
    Dim childLevel As Long
    Dim i As Long
    Dim durationVal As Variant
    Dim progressVal As Variant
    Dim totalDuration As Double
    Dim weightedProgress As Double
    Dim childId As String

    If Not mapTbl.Exists("WBS") Then Exit Function
    If Not mapTbl.Exists("Calculated Duration") Then Exit Function

    parentWbs = CStr(Dashboard_ArrayVal(arr, mapTbl, r, "WBS"))
    parentLevel = Dashboard_WbsLevel(parentWbs)

    For i = 1 To UBound(arr, 1)
        childWbs = CStr(Dashboard_ArrayVal(arr, mapTbl, i, "WBS"))
        childLevel = Dashboard_WbsLevel(childWbs)
        If childLevel = parentLevel + 1 Then
            If Left$(childWbs, Len(parentWbs) + 1) = parentWbs & "." Then
                durationVal = GetCellValue(Dashboard_ArrayVal(arr, mapTbl, i, "Calculated Duration"))
                If IsNumeric(durationVal) And CDbl(durationVal) > 0 Then
                    childId = CStr(GetCellValue(Dashboard_ArrayVal(arr, mapTbl, i, "ID")))
                    progressVal = Dashboard_RowProgress(arr, mapTbl, progressById, i, childId)
                    totalDuration = totalDuration + CDbl(durationVal)
                    weightedProgress = weightedProgress + (CDbl(durationVal) * CDbl(progressVal))
                End If
            End If
        End If
    Next i

    If totalDuration > 0 Then
        Dashboard_SummaryProgress = weightedProgress / totalDuration
    Else
        Dashboard_SummaryProgress = Dashboard_RowProgress(arr, mapTbl, progressById, r, CStr(GetCellValue(Dashboard_ArrayVal(arr, mapTbl, r, "ID"))))
    End If

End Function

'------------------------------------------------------------------------------
' FR: Retourne la map Row Progress sans modifier les donnees d'entree.
' EN: Returns the Row Progress map without mutating input data.
'------------------------------------------------------------------------------

Private Function Dashboard_RowProgress(ByRef arr As Variant, ByVal mapTbl As Object, ByVal progressById As Object, ByVal r As Long, ByVal idVal As String) As Double

    Dim actualFinish As Variant

    If Not progressById Is Nothing Then
        If idVal <> "" And progressById.Exists(idVal) Then
            Dashboard_RowProgress = CDbl(progressById(idVal))
            Exit Function
        End If
    End If

    If mapTbl.Exists("% Progress") Then
        Dashboard_RowProgress = Dashboard_NormalizeProgress(GetCellValue(Dashboard_ArrayVal(arr, mapTbl, r, "% Progress")))
    ElseIf mapTbl.Exists("Actual Finish") Then
        actualFinish = GetCellValue(Dashboard_ArrayVal(arr, mapTbl, r, "Actual Finish"))
        If Dashboard_HasDateValue(actualFinish) Then Dashboard_RowProgress = 1#
    End If

End Function

'------------------------------------------------------------------------------
' FR: Normalise ou formate Normalize Progress selon le contrat canonique du composant.
' EN: Normalizes or formats Normalize Progress according to the component contract.
'------------------------------------------------------------------------------

Private Function Dashboard_NormalizeProgress(ByVal progressVal As Variant) As Double

    If Not HasValue(progressVal) Or Not IsNumeric(progressVal) Then Exit Function

    Dashboard_NormalizeProgress = CDbl(progressVal)
    If Dashboard_NormalizeProgress > 1# Then Dashboard_NormalizeProgress = Dashboard_NormalizeProgress / 100#
    If Dashboard_NormalizeProgress < 0# Then Dashboard_NormalizeProgress = 0#
    If Dashboard_NormalizeProgress > 1# Then Dashboard_NormalizeProgress = 1#

End Function

'------------------------------------------------------------------------------
' FR: Retourne la valeur Wbs Level sans modifier les donnees d'entree.
' EN: Returns the Wbs Level value without mutating input data.
'------------------------------------------------------------------------------

Private Function Dashboard_WbsLevel(ByVal wbsValue As String) As Long

    Dim txt As String
    Dim i As Long

    txt = Replace$(Trim$(wbsValue), ",", ".")
    If txt = "" Then Exit Function

    Dashboard_WbsLevel = 1
    For i = 1 To Len(txt)
        If Mid$(txt, i, 1) = "." Then Dashboard_WbsLevel = Dashboard_WbsLevel + 1
    Next i

End Function

'------------------------------------------------------------------------------
' FR: Retourne la valeur Date To X sans modifier les donnees d'entree.
' EN: Returns the Date To X value without mutating input data.
'------------------------------------------------------------------------------

Private Function Dashboard_DateToX(ByVal dateVal As Double, ByVal minDate As Double, ByVal maxDate As Double, ByVal leftPos As Double, ByVal widthVal As Double) As Double

    If maxDate <= minDate Then
        Dashboard_DateToX = leftPos
    Else
        Dashboard_DateToX = leftPos + (((dateVal - minDate) / (maxDate - minDate)) * widthVal)
    End If

End Function

'------------------------------------------------------------------------------
' FR: Retourne la collection Planning Overview Rows sans modifier les donnees d'entree.
' EN: Returns the Planning Overview Rows collection without mutating input data.
'------------------------------------------------------------------------------

Private Function Dashboard_PlanningOverviewRows(ByRef arr As Variant, ByVal mapTbl As Object, ByRef minOut As Double, ByRef maxOut As Double) As Collection

    Const MAX_OVERVIEW_ROWS As Long = 10

    Dim activeRows As Collection
    Dim beforeRows As Collection
    Dim afterRows As Collection
    Dim selected As Collection
    Dim selectedRows As Object
    Dim r As Long
    Dim rowStart As Variant
    Dim rowFinish As Variant
    Dim startSerial As Double
    Dim finishSerial As Double
    Dim todaySerial As Double
    Dim remainingSlots As Long
    Dim beforeQuota As Long
    Dim afterQuota As Long
    Dim hasDate As Boolean
    Dim orderedRows As Collection

    Set activeRows = New Collection
    Set beforeRows = New Collection
    Set afterRows = New Collection
    Set selected = New Collection
    Set selectedRows = CreateObject("Scripting.Dictionary")

    todaySerial = CDbl(Date)

    For r = 1 To UBound(arr, 1)
        If Dashboard_IsPlanningOverviewVisibleRow(arr, mapTbl, r) Then
            If Dashboard_PlanningRowHasRenderableDates(arr, mapTbl, r) Then
                rowStart = GetCellValue(Dashboard_ArrayVal(arr, mapTbl, r, "Calculated Start"))
                rowFinish = GetCellValue(Dashboard_ArrayVal(arr, mapTbl, r, "Calculated Finish"))
                startSerial = Dashboard_DateNumber(rowStart)
                finishSerial = Dashboard_DateNumber(rowFinish)

                If startSerial <= todaySerial And finishSerial >= todaySerial Then
                    activeRows.Add r
                ElseIf finishSerial < todaySerial Then
                    Dashboard_InsertPlanningCandidateByDistance beforeRows, r, todaySerial - finishSerial
                ElseIf startSerial > todaySerial Then
                    Dashboard_InsertPlanningCandidateByDistance afterRows, r, startSerial - todaySerial
                End If
            End If
        End If
    Next r

    Dashboard_AddPlanningCandidates selected, selectedRows, activeRows, MAX_OVERVIEW_ROWS

    remainingSlots = MAX_OVERVIEW_ROWS - selected.Count
    If remainingSlots > 0 Then
        beforeQuota = remainingSlots \ 2
        afterQuota = remainingSlots - beforeQuota

        Dashboard_AddPlanningCandidates selected, selectedRows, beforeRows, beforeQuota
        Dashboard_AddPlanningCandidates selected, selectedRows, afterRows, afterQuota

        If selected.Count < MAX_OVERVIEW_ROWS Then
            Dashboard_AddPlanningCandidates selected, selectedRows, beforeRows, MAX_OVERVIEW_ROWS - selected.Count
        End If
        If selected.Count < MAX_OVERVIEW_ROWS Then
            Dashboard_AddPlanningCandidates selected, selectedRows, afterRows, MAX_OVERVIEW_ROWS - selected.Count
        End If
    End If

    Set orderedRows = Dashboard_OrderPlanningRowsBySourceOrder(selected)

    For r = 1 To orderedRows.Count
        Dashboard_IncludeAxisDate arr, mapTbl, CLng(orderedRows(r)), "Baseline Start", hasDate, minOut, maxOut
        Dashboard_IncludeAxisDate arr, mapTbl, CLng(orderedRows(r)), "Baseline Finish", hasDate, minOut, maxOut
        Dashboard_IncludeAxisDate arr, mapTbl, CLng(orderedRows(r)), "Calculated Start", hasDate, minOut, maxOut
        Dashboard_IncludeAxisDate arr, mapTbl, CLng(orderedRows(r)), "Calculated Finish", hasDate, minOut, maxOut
    Next r

    Set Dashboard_PlanningOverviewRows = orderedRows

End Function

'------------------------------------------------------------------------------
' FR: Traite la collection Insert Planning Candidate By Distance sans modifier les donnees d'entree.
' EN: Handles the Insert Planning Candidate By Distance collection without mutating input data.
'------------------------------------------------------------------------------

Private Sub Dashboard_InsertPlanningCandidateByDistance(ByVal rows As Collection, ByVal rowIndex As Long, ByVal distanceVal As Double)

    Dim item As Object
    Dim existing As Object
    Dim tmp As Collection
    Dim inserted As Boolean
    Dim i As Long

    Set item = CreateObject("Scripting.Dictionary")
    item("Row") = rowIndex
    item("Distance") = distanceVal

    Set tmp = New Collection
    For i = 1 To rows.Count
        Set existing = rows(i)
        If Not inserted Then
            If distanceVal < CDbl(existing("Distance")) Then
                tmp.Add item
                inserted = True
            End If
        End If
        tmp.Add existing
    Next i

    If Not inserted Then tmp.Add item

    Do While rows.Count > 0
        rows.Remove 1
    Loop
    For i = 1 To tmp.Count
        rows.Add tmp(i)
    Next i

End Sub

'------------------------------------------------------------------------------
' FR: Ajoute la collection Planning Candidates a la structure cible fournie par l'appelant.
' EN: Adds the Planning Candidates collection to the target structure supplied by the caller.
'------------------------------------------------------------------------------

Private Sub Dashboard_AddPlanningCandidates( _
    ByVal selected As Collection, _
    ByVal selectedRows As Object, _
    ByVal candidates As Collection, _
    ByVal maxToAdd As Long)

    Dim item As Variant
    Dim rowIndex As Long
    Dim added As Long

    If maxToAdd <= 0 Then Exit Sub

    For Each item In candidates
        rowIndex = Dashboard_PlanningCandidateRow(item)
        If rowIndex > 0 Then
            If Not selectedRows.Exists(CStr(rowIndex)) Then
                selectedRows(CStr(rowIndex)) = True
                selected.Add rowIndex
                added = added + 1
                If added >= maxToAdd Then Exit Sub
            End If
        End If
    Next item

End Sub

'------------------------------------------------------------------------------
' FR: Retourne la valeur Planning Candidate Row sans modifier les donnees d'entree.
' EN: Returns the Planning Candidate Row value without mutating input data.
'------------------------------------------------------------------------------

Private Function Dashboard_PlanningCandidateRow(ByVal item As Variant) As Long

    If IsObject(item) Then
        Dashboard_PlanningCandidateRow = CLng(item("Row"))
    Else
        Dashboard_PlanningCandidateRow = CLng(item)
    End If

End Function

'------------------------------------------------------------------------------
' FR: Retourne la collection Order Planning Rows By Source Order sans modifier les donnees d'entree.
' EN: Returns the Order Planning Rows By Source Order collection without mutating input data.
'------------------------------------------------------------------------------

Private Function Dashboard_OrderPlanningRowsBySourceOrder(ByVal rows As Collection) As Collection

    Dim ordered As Collection
    Dim i As Long
    Dim j As Long
    Dim currentRow As Long
    Dim inserted As Boolean

    Set ordered = New Collection

    For i = 1 To rows.Count
        currentRow = CLng(rows(i))
        inserted = False

        For j = 1 To ordered.Count
            If currentRow < CLng(ordered(j)) Then
                ordered.Add currentRow, , j
                inserted = True
                Exit For
            End If
        Next j

        If Not inserted Then ordered.Add currentRow
    Next i

    Set Dashboard_OrderPlanningRowsBySourceOrder = ordered

End Function
'------------------------------------------------------------------------------
' FR: Indique si la map Planning Overview Visible Row satisfait la condition attendue, sans modifier les donnees source.
' EN: Returns whether the Planning Overview Visible Row map satisfies the expected condition without mutating source data.
'------------------------------------------------------------------------------

Private Function Dashboard_IsPlanningOverviewVisibleRow(ByRef arr As Variant, ByVal mapTbl As Object, ByVal r As Long) As Boolean

    Dim durationVal As Variant
    Dim summaryDisplayVal As String

    If mapTbl.Exists("S") Then
        summaryDisplayVal = UCase$(Trim$(CStr(Dashboard_ArrayVal(arr, mapTbl, r, "S"))))
        Dashboard_IsPlanningOverviewVisibleRow = (summaryDisplayVal = "Y")
        Exit Function
    End If

    If Dashboard_IsSummaryArrayRow(arr, mapTbl, r) Then
        Dashboard_IsPlanningOverviewVisibleRow = True
        Exit Function
    End If

    If mapTbl.Exists("Calculated Duration") Then
        durationVal = GetCellValue(Dashboard_ArrayVal(arr, mapTbl, r, "Calculated Duration"))
        If IsNumeric(durationVal) Then
            Dashboard_IsPlanningOverviewVisibleRow = (CDbl(durationVal) <= 1#)
        End If
    End If

End Function

'------------------------------------------------------------------------------
' FR: Retourne la map Planning Row Has Renderable Dates sans modifier les donnees d'entree.
' EN: Returns the Planning Row Has Renderable Dates map without mutating input data.
'------------------------------------------------------------------------------

Private Function Dashboard_PlanningRowHasRenderableDates(ByRef arr As Variant, ByVal mapTbl As Object, ByVal r As Long) As Boolean

    Dim startVal As Variant
    Dim finishVal As Variant

    startVal = GetCellValue(Dashboard_ArrayVal(arr, mapTbl, r, "Calculated Start"))
    finishVal = GetCellValue(Dashboard_ArrayVal(arr, mapTbl, r, "Calculated Finish"))
    Dashboard_PlanningRowHasRenderableDates = (Dashboard_HasDateValue(startVal) And Dashboard_HasDateValue(finishVal))

End Function

'------------------------------------------------------------------------------
' FR: Retourne la valeur Planning Axis Margin sans modifier les donnees d'entree.
' EN: Returns the Planning Axis Margin value without mutating input data.
'------------------------------------------------------------------------------

Private Function Dashboard_PlanningAxisMargin(ByVal minDate As Double, ByVal maxDate As Double) As Double

    Dim spanDays As Double

    spanDays = maxDate - minDate
    If spanDays <= 14 Then
        Dashboard_PlanningAxisMargin = 1
    ElseIf spanDays <= 62 Then
        Dashboard_PlanningAxisMargin = 3
    ElseIf spanDays <= 730 Then
        Dashboard_PlanningAxisMargin = 14
    Else
        Dashboard_PlanningAxisMargin = 45
    End If

End Function

'------------------------------------------------------------------------------
' FR: Retourne la map Planning Axis Bounds sans exposer de mutateur sur l'etat source.
' EN: Returns the Planning Axis Bounds map without exposing a mutator for source state.
'------------------------------------------------------------------------------

Private Sub Dashboard_GetPlanningAxisBounds(ByRef arr As Variant, ByVal mapTbl As Object, ByRef minOut As Double, ByRef maxOut As Double)

    Dim r As Long
    Dim shown As Long
    Dim v As Variant
    Dim hasDate As Boolean

    For r = 1 To UBound(arr, 1)
        If Dashboard_IsSummaryArrayRow(arr, mapTbl, r) Then
            If Dashboard_WbsLevel(CStr(Dashboard_ArrayVal(arr, mapTbl, r, "WBS"))) <= 3 Then
                shown = shown + 1
                If shown > 7 Then Exit For
                Dashboard_IncludeAxisDate arr, mapTbl, r, "Baseline Start", hasDate, minOut, maxOut
                Dashboard_IncludeAxisDate arr, mapTbl, r, "Baseline Finish", hasDate, minOut, maxOut
                Dashboard_IncludeAxisDate arr, mapTbl, r, "Calculated Start", hasDate, minOut, maxOut
                Dashboard_IncludeAxisDate arr, mapTbl, r, "Calculated Finish", hasDate, minOut, maxOut
            End If
        End If
    Next r

    If hasDate Then
        minOut = minOut - 7
        maxOut = maxOut + 10
    End If

End Sub

'------------------------------------------------------------------------------
' FR: Traite la map Include Axis Date sans modifier les donnees d'entree.
' EN: Handles the Include Axis Date map without mutating input data.
'------------------------------------------------------------------------------

Private Sub Dashboard_IncludeAxisDate(ByRef arr As Variant, ByVal mapTbl As Object, ByVal r As Long, ByVal colName As String, ByRef hasDate As Boolean, ByRef minOut As Double, ByRef maxOut As Double)

    Dim v As Variant
    Dim d As Double

    v = GetCellValue(Dashboard_ArrayVal(arr, mapTbl, r, colName))
    If Not Dashboard_HasDateValue(v) Then Exit Sub

    d = Dashboard_DateNumber(v)
    If Not hasDate Then
        minOut = d
        maxOut = d
        hasDate = True
    ElseIf d < minOut Then
        minOut = d
    ElseIf d > maxOut Then
        maxOut = d
    End If

End Sub

'------------------------------------------------------------------------------
' FR: Insere un candidat dans le classement borne des plus forts retards du Dashboard.
' EN: Inserts a candidate into the Dashboard's bounded highest-delay ranking.
'------------------------------------------------------------------------------

Private Sub Dashboard_InsertTopDelay( _
    ByRef topIds() As String, _
    ByRef topWbs() As String, _
    ByRef topNames() As String, _
    ByRef topVals() As Double, _
    ByVal idVal As String, _
    ByVal wbsVal As String, _
    ByVal nameVal As String, _
    ByVal metricVal As Double)

    Dim i As Long
    Dim j As Long

    For i = 1 To 5
        If metricVal > topVals(i) Then
            For j = 5 To i + 1 Step -1
                topIds(j) = topIds(j - 1)
                topWbs(j) = topWbs(j - 1)
                topNames(j) = topNames(j - 1)
                topVals(j) = topVals(j - 1)
            Next j
            topIds(i) = idVal
            topWbs(i) = wbsVal
            topNames(i) = nameVal
            topVals(i) = metricVal
            Exit Sub
        End If
    Next i

End Sub

'------------------------------------------------------------------------------
' FR: Insere un candidat dans le classement borne des risques les plus faibles du Dashboard.
' EN: Inserts a candidate into the Dashboard's bounded lowest-risk ranking.
'------------------------------------------------------------------------------

Private Sub Dashboard_InsertLowestRisk( _
    ByRef ids() As String, _
    ByRef wbsVals() As String, _
    ByRef names() As String, _
    ByRef vals() As Double, _
    ByVal idVal As String, _
    ByVal wbsVal As String, _
    ByVal nameVal As String, _
    ByVal metricVal As Double)

    Dim i As Long
    Dim j As Long

    For i = 1 To 5
        If metricVal < vals(i) Then
            For j = 5 To i + 1 Step -1
                ids(j) = ids(j - 1)
                wbsVals(j) = wbsVals(j - 1)
                names(j) = names(j - 1)
                vals(j) = vals(j - 1)
            Next j
            ids(i) = idVal
            wbsVals(i) = wbsVal
            names(i) = nameVal
            vals(i) = metricVal
            Exit Sub
        End If
    Next i

End Sub

'------------------------------------------------------------------------------
' FR: Actualise Update Kpi Card sans modifier les regles metier qui produisent les donnees.
' EN: Refreshes Update Kpi Card without changing the business rules that produce the data.
' FR - Effet de bord : cree ou met a jour des shapes Excel.
' EN - Side effect: creates or updates Excel shapes.
'------------------------------------------------------------------------------

Private Sub Dashboard_UpdateKpiCard(ByVal ws As Worksheet, ByVal kpiIndex As Long, ByVal titleText As String, ByVal valueText As String, ByVal subText As String, ByVal x As Double, ByVal y As Double, ByVal w As Double, ByVal h As Double, ByVal accentColor As Long)

    Dim bg As Shape
    Dim accent As Shape
    Dim titleBox As Shape
    Dim valueBox As Shape
    Dim subBox As Shape
    Dim prefixName As String

    prefixName = DASH_PREFIX & "KPI_" & CStr(kpiIndex) & "_"

    Set bg = Dashboard_GetShape(ws, prefixName & "BG")
    Set accent = Dashboard_GetShape(ws, prefixName & "ACCENT")
    Set titleBox = Dashboard_GetShape(ws, prefixName & "TITLE")
    Set valueBox = Dashboard_GetShape(ws, prefixName & "HERO")
    Set subBox = Dashboard_GetShape(ws, prefixName & "SUB")

    If bg Is Nothing Or accent Is Nothing Or titleBox Is Nothing Or valueBox Is Nothing Or subBox Is Nothing Then
        Dashboard_DeleteShapesInRect ws, x - 2, y - 2, w + 4, h + 4

        Set bg = ws.Shapes.AddShape(msoShapeRoundedRectangle, x, y, w, h)
        bg.Name = prefixName & "BG"
        bg.Fill.ForeColor.RGB = RGB(255, 255, 255)
        bg.Line.ForeColor.RGB = RGB(224, 230, 237)

        Set accent = ws.Shapes.AddShape(msoShapeRectangle, x, y, 6, h)
        accent.Name = prefixName & "ACCENT"
        accent.Line.Visible = msoFalse

        Set titleBox = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, x + 16, y + 8, w - 24, 16)
        titleBox.Name = prefixName & "TITLE"
        Dashboard_FormatStableTextBox titleBox, 9, RGB(96, 111, 128), False

        Set valueBox = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, x + 16, y + 24, w - 24, 34)
        valueBox.Name = prefixName & "HERO"
        Dashboard_FormatStableTextBox valueBox, 21, RGB(20, 34, 51), True

        Set subBox = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, x + 16, y + h - 20, w - 24, 18)
        subBox.Name = prefixName & "SUB"
        Dashboard_FormatStableTextBox subBox, 8, RGB(96, 111, 128), False
    End If

    Dashboard_PositionKpiShape bg, x, y, w, h
    Dashboard_PositionKpiShape accent, x, y, 6, h
    Dashboard_PositionKpiShape titleBox, x + 16, y + 8, w - 24, 16
    Dashboard_PositionKpiShape valueBox, x + 16, y + 24, w - 24, 34
    Dashboard_PositionKpiShape subBox, x + 16, y + h - 20, w - 24, 18

    bg.Fill.ForeColor.RGB = RGB(255, 255, 255)
    bg.Line.ForeColor.RGB = RGB(224, 230, 237)
    accent.Fill.ForeColor.RGB = accentColor
    accent.Line.Visible = msoFalse

    Dashboard_SetStableText titleBox, titleText, 9, RGB(96, 111, 128), False
    Dashboard_SetStableText valueBox, valueText, 21, RGB(20, 34, 51), True
    Dashboard_SetStableText subBox, subText, 8, RGB(96, 111, 128), False

End Sub

'------------------------------------------------------------------------------
' FR: Retourne la reference Shape sans exposer de mutateur sur l'etat source.
' EN: Returns the Shape reference without exposing a mutator for source state.
' FR - Effet de bord : cree ou met a jour des shapes Excel.
' EN - Side effect: creates or updates Excel shapes.
'------------------------------------------------------------------------------

Private Function Dashboard_GetShape(ByVal ws As Worksheet, ByVal shapeName As String) As Shape

    On Error Resume Next
    Set Dashboard_GetShape = ws.Shapes(shapeName)
    On Error GoTo 0

End Function

'------------------------------------------------------------------------------
' FR: Traite la reference Position Kpi Shape sans modifier les donnees d'entree.
' EN: Handles the Position Kpi Shape reference without mutating input data.
'------------------------------------------------------------------------------

Private Sub Dashboard_PositionKpiShape(ByVal shp As Shape, ByVal x As Double, ByVal y As Double, ByVal w As Double, ByVal h As Double)

    shp.Left = x
    shp.Top = y
    shp.Width = w
    shp.Height = h

End Sub

'------------------------------------------------------------------------------
' FR: Reinitialise Delete Shapes In Rect dans le perimetre possede par le composant.
' EN: Resets Delete Shapes In Rect within the state owned by the component.
' FR - Effet de bord : efface uniquement les donnees ou objets cibles du contrat.
' EN - Side effect: clears only data or objects targeted by the contract.
'------------------------------------------------------------------------------

Private Sub Dashboard_DeleteShapesInRect(ByVal ws As Worksheet, ByVal x As Double, ByVal y As Double, ByVal w As Double, ByVal h As Double)

    Dim i As Long
    Dim shp As Shape
    Dim rightEdge As Double
    Dim bottomEdge As Double

    rightEdge = x + w
    bottomEdge = y + h

    For i = ws.Shapes.Count To 1 Step -1
        Set shp = ws.Shapes(i)
        If shp.Left >= x And shp.Left <= rightEdge And shp.Top >= y And shp.Top <= bottomEdge Then
            shp.Delete
        End If
    Next i

End Sub

'------------------------------------------------------------------------------
' FR: Normalise ou formate Format Stable Text Box selon le contrat canonique du composant.
' EN: Normalizes or formats Format Stable Text Box according to the component contract.
'------------------------------------------------------------------------------

Private Sub Dashboard_FormatStableTextBox(ByVal shp As Shape, ByVal fontSize As Double, ByVal colorVal As Long, ByVal isBold As Boolean)

    shp.Fill.Visible = msoFalse
    shp.Line.Visible = msoFalse
    shp.TextFrame2.MarginLeft = 0
    shp.TextFrame2.MarginRight = 0
    shp.TextFrame2.MarginTop = 0
    shp.TextFrame2.MarginBottom = 0
    shp.TextFrame2.VerticalAnchor = msoAnchorMiddle
    shp.TextFrame2.WordWrap = msoFalse
    shp.TextFrame2.TextRange.Font.Name = "Segoe UI"
    shp.TextFrame2.TextRange.Font.Size = fontSize
    shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = colorVal
    shp.TextFrame2.TextRange.Font.Bold = isBold

End Sub

'------------------------------------------------------------------------------
' FR: Active ou initialise Set Stable Text dans l'etat runtime du composant.
' EN: Activates or initializes Set Stable Text in the component runtime state.
'------------------------------------------------------------------------------

Private Sub Dashboard_SetStableText(ByVal shp As Shape, ByVal txt As String, ByVal fontSize As Double, ByVal colorVal As Long, ByVal isBold As Boolean)

    shp.TextFrame2.TextRange.Text = txt
    Dashboard_FormatStableTextBox shp, fontSize, colorVal, isBold

End Sub
'------------------------------------------------------------------------------
' FR: Ajoute la reference Panel a la structure cible fournie par l'appelant.
' EN: Adds the Panel reference to the target structure supplied by the caller.
' FR - Effet de bord : cree ou met a jour des shapes Excel.
' EN - Side effect: creates or updates Excel shapes.
'------------------------------------------------------------------------------

Private Sub Dashboard_AddPanel(ByVal ws As Worksheet, ByVal x As Double, ByVal y As Double, ByVal w As Double, ByVal h As Double, ByVal fillColor As Long)

    Dim shp As Shape

    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, x, y, w, h)
    shp.Name = DASH_PREFIX & "PANEL"
    shp.Fill.ForeColor.RGB = fillColor
    shp.Line.ForeColor.RGB = RGB(224, 230, 237)

End Sub

'------------------------------------------------------------------------------
' FR: Ajoute la reference Slide Surface a la structure cible fournie par l'appelant.
' EN: Adds the Slide Surface reference to the target structure supplied by the caller.
' FR - Effet de bord : cree ou met a jour des shapes Excel.
' EN - Side effect: creates or updates Excel shapes.
'------------------------------------------------------------------------------

Private Sub Dashboard_AddSlideSurface(ByVal ws As Worksheet, ByVal x As Double, ByVal y As Double, ByVal w As Double, ByVal h As Double)

    Dim shp As Shape

    Set shp = ws.Shapes.AddShape(msoShapeRectangle, x, y, w, h)
    shp.Name = DASH_PREFIX & "SLIDE"
    shp.Fill.ForeColor.RGB = RGB(255, 255, 255)
    shp.Line.ForeColor.RGB = RGB(224, 230, 237)
    shp.ZOrder msoSendToBack

End Sub

'------------------------------------------------------------------------------
' FR: Ajoute la reference Section Title a la structure cible fournie par l'appelant.
' EN: Adds the Section Title reference to the target structure supplied by the caller.
' FR - Effet de bord : cree ou met a jour des shapes Excel.
' EN - Side effect: creates or updates Excel shapes.
'------------------------------------------------------------------------------

Private Sub Dashboard_AddSectionTitle(ByVal ws As Worksheet, ByVal titleText As String, ByVal x As Double, ByVal y As Double, ByVal w As Double)

    Dim shp As Shape

    Set shp = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, x, y, w, 20)
    Dashboard_FormatTextBox shp, titleText, 12, RGB(20, 34, 51), True

End Sub

'------------------------------------------------------------------------------
' FR: Ajoute la reference Centered Section Title a la structure cible fournie par l'appelant.
' EN: Adds the Centered Section Title reference to the target structure supplied by the caller.
' FR - Effet de bord : cree ou met a jour des shapes Excel.
' EN - Side effect: creates or updates Excel shapes.
'------------------------------------------------------------------------------

Private Sub Dashboard_AddCenteredSectionTitle(ByVal ws As Worksheet, ByVal titleText As String, ByVal x As Double, ByVal y As Double, ByVal w As Double)

    Dim shp As Shape

    Set shp = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, x, y, w, 20)
    Dashboard_FormatTextBox shp, titleText, 12, RGB(20, 34, 51), True
    shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter

End Sub

'------------------------------------------------------------------------------
' FR: Ajoute la reference Table Title a la structure cible fournie par l'appelant.
' EN: Adds the Table Title reference to the target structure supplied by the caller.
' FR - Effet de bord : cree ou met a jour des shapes Excel.
' EN - Side effect: creates or updates Excel shapes.
'------------------------------------------------------------------------------

Private Sub Dashboard_AddTableTitle(ByVal ws As Worksheet, ByVal titleText As String, ByVal x As Double, ByVal y As Double, ByVal w As Double)

    Dim shp As Shape

    Set shp = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, x, y, w, 18)
    Dashboard_FormatTextBox shp, titleText, 11, RGB(20, 34, 51), True

End Sub

'------------------------------------------------------------------------------
' FR: Ecrit ou synchronise Write Hotspot Header dans le stockage possede par le domaine.
' EN: Writes or synchronizes Write Hotspot Header in the store owned by the domain.
'------------------------------------------------------------------------------

Private Sub Dashboard_WriteHotspotHeader(ByVal ws As Worksheet, ByVal x As Double, ByVal y As Double, ByVal headers As Variant)

    Dashboard_AddTinyText ws, CStr(headers(0)), x, y, 54, 14, RGB(96, 111, 128)
    Dashboard_AddTinyText ws, CStr(headers(1)), x + 58, y, 138, 14, RGB(96, 111, 128)
    Dashboard_AddTinyText ws, CStr(headers(2)), x + 202, y, 52, 14, RGB(96, 111, 128)

End Sub

'------------------------------------------------------------------------------
' FR: Ecrit ou synchronise Write Hotspot Row dans le stockage possede par le domaine.
' EN: Writes or synchronizes Write Hotspot Row in the store owned by the domain.
'------------------------------------------------------------------------------

Private Sub Dashboard_WriteHotspotRow(ByVal ws As Worksheet, ByVal x As Double, ByVal y As Double, ByVal wbsVal As String, ByVal nameVal As String, ByVal metricText As String, ByVal metricColor As Long)

    Dashboard_AddTinyText ws, wbsVal, x, y, 54, 16, RGB(20, 34, 51)
    Dashboard_AddTinyText ws, nameVal, x + 58, y, 138, 16, RGB(42, 52, 65)
    Dashboard_AddTinyText ws, metricText, x + 202, y, 52, 16, metricColor

End Sub

'------------------------------------------------------------------------------
' FR: Ajoute la reference Mini Gantt Label a la structure cible fournie par l'appelant.
' EN: Adds the Mini Gantt Label reference to the target structure supplied by the caller.
' FR - Effet de bord : cree ou met a jour des shapes Excel.
' EN - Side effect: creates or updates Excel shapes.
'------------------------------------------------------------------------------

Private Sub Dashboard_AddMiniGanttLabel(ByVal ws As Worksheet, ByVal txt As String, ByVal x As Double, ByVal y As Double, ByVal w As Double, ByVal fontSize As Double)

    Dim shp As Shape

    Set shp = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, x, y, w, fontSize + 4)
    Dashboard_FormatTextBox shp, txt, fontSize, RGB(42, 52, 65), False
    shp.TextFrame2.VerticalAnchor = msoAnchorMiddle
    shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignLeft

End Sub

'------------------------------------------------------------------------------
' FR: Retourne la valeur Truncate Text sans modifier les donnees d'entree.
' EN: Returns the Truncate Text value without mutating input data.
'------------------------------------------------------------------------------

Private Function Dashboard_TruncateText(ByVal txt As String, ByVal maxLen As Long) As String

    If Len(txt) <= maxLen Then
        Dashboard_TruncateText = txt
    ElseIf maxLen <= 3 Then
        Dashboard_TruncateText = Left$(txt, maxLen)
    Else
        Dashboard_TruncateText = Left$(txt, maxLen - 3) & "..."
    End If

End Function

'------------------------------------------------------------------------------
' FR: Ajoute la reference Tiny Label a la structure cible fournie par l'appelant.
' EN: Adds the Tiny Label reference to the target structure supplied by the caller.
'------------------------------------------------------------------------------

Private Sub Dashboard_AddTinyLabel(ByVal ws As Worksheet, ByVal txt As String, ByVal x As Double, ByVal y As Double, ByVal w As Double)

    Dashboard_AddTinyText ws, txt, x, y, w, 14, RGB(96, 111, 128)

End Sub

'------------------------------------------------------------------------------
' FR: Ajoute la reference Tiny Text a la structure cible fournie par l'appelant.
' EN: Adds the Tiny Text reference to the target structure supplied by the caller.
' FR - Effet de bord : cree ou met a jour des shapes Excel.
' EN - Side effect: creates or updates Excel shapes.
'------------------------------------------------------------------------------

Private Sub Dashboard_AddTinyText(ByVal ws As Worksheet, ByVal txt As String, ByVal x As Double, ByVal y As Double, ByVal w As Double, ByVal h As Double, ByVal colorVal As Long)

    Dim shp As Shape

    Set shp = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, x, y, w, h)
    Dashboard_FormatTextBox shp, txt, 8, colorVal, False

End Sub

'------------------------------------------------------------------------------
' FR: Ajoute la reference Axis Tick Label a la structure cible fournie par l'appelant.
' EN: Adds the Axis Tick Label reference to the target structure supplied by the caller.
' FR - Effet de bord : cree ou met a jour des shapes Excel.
' EN - Side effect: creates or updates Excel shapes.
'------------------------------------------------------------------------------

Private Sub Dashboard_AddAxisTickLabel(ByVal ws As Worksheet, ByVal txt As String, ByVal x As Double, ByVal y As Double, ByVal w As Double, ByVal h As Double, ByVal colorVal As Long)

    Dim shp As Shape

    Set shp = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, x, y, w, h)
    Dashboard_FormatTextBox shp, txt, 8, colorVal, False
    shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter

End Sub

'------------------------------------------------------------------------------
' FR: Ecrit ou synchronise Write Empty State dans le stockage possede par le domaine.
' EN: Writes or synchronizes Write Empty State in the store owned by the domain.
'------------------------------------------------------------------------------

Private Sub Dashboard_WriteEmptyState(ByVal ws As Worksheet, ByVal x As Double, ByVal y As Double, ByVal w As Double, ByVal txt As String)

    Dashboard_AddTinyText ws, Dashboard_L("OK - ", "OK - ") & txt, x, y, w, 16, RGB(0, 145, 112)

End Sub

'------------------------------------------------------------------------------
' FR: Ecrit ou synchronise Write Dashboard Empty State dans le stockage possede par le domaine.
' EN: Writes or synchronizes Write Dashboard Empty State in the store owned by the domain.
' FR - Effet de bord : cree ou met a jour des shapes Excel.
' EN - Side effect: creates or updates Excel shapes.
'------------------------------------------------------------------------------

Private Sub Dashboard_WriteDashboardEmptyState(ByVal ws As Worksheet, ByVal x As Double, ByVal y As Double, ByVal w As Double)

    Dim shp As Shape

    Set shp = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, x, y, w, 18)
    Dashboard_FormatTextBox shp, Dashboard_NoProjectLoadedText(), 9, RGB(96, 111, 128), True
    shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter

End Sub

'------------------------------------------------------------------------------
' FR: Ajoute la reference Badge a la structure cible fournie par l'appelant.
' EN: Adds the Badge reference to the target structure supplied by the caller.
' FR - Effet de bord : cree ou met a jour des shapes Excel.
' EN - Side effect: creates or updates Excel shapes.
'------------------------------------------------------------------------------

Private Sub Dashboard_AddBadge(ByVal ws As Worksheet, ByVal txt As String, ByVal x As Double, ByVal y As Double, ByVal w As Double, ByVal h As Double, ByVal fillColor As Long, ByVal textColor As Long)

    Dim bg As Shape
    Dim t As Shape

    Set bg = ws.Shapes.AddShape(msoShapeRoundedRectangle, x, y, w, h)
    bg.Name = DASH_PREFIX & "BADGE"
    bg.Fill.ForeColor.RGB = fillColor
    bg.Line.Visible = msoFalse

    Set t = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, x + 8, y + 3, w - 16, h - 6)
    Dashboard_FormatTextBox t, txt, 8, textColor, True
    t.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter

End Sub

'------------------------------------------------------------------------------
' FR: Ajoute la reference Bar a la structure cible fournie par l'appelant.
' EN: Adds the Bar reference to the target structure supplied by the caller.
' FR - Effet de bord : cree ou met a jour des shapes Excel.
' EN - Side effect: creates or updates Excel shapes.
'------------------------------------------------------------------------------

Private Sub Dashboard_AddBar(ByVal ws As Worksheet, ByVal x As Double, ByVal y As Double, ByVal w As Double, ByVal h As Double, ByVal colorVal As Long)

    Dim shp As Shape

    If w < 2 Then w = 2
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, x, y, w, h)
    shp.Name = DASH_PREFIX & "BAR"
    shp.Fill.ForeColor.RGB = colorVal
    shp.Line.Visible = msoFalse

End Sub

'------------------------------------------------------------------------------
' FR: Ajoute la reference Milestone Diamond a la structure cible fournie par l'appelant.
' EN: Adds the Milestone Diamond reference to the target structure supplied by the caller.
' FR - Effet de bord : cree ou met a jour des shapes Excel.
' EN - Side effect: creates or updates Excel shapes.
'------------------------------------------------------------------------------

Private Sub Dashboard_AddMilestoneDiamond(ByVal ws As Worksheet, ByVal centerX As Double, ByVal centerY As Double, ByVal sizeVal As Double, ByVal colorVal As Long)

    Dim shp As Shape

    If sizeVal < 7 Then sizeVal = 7
    Set shp = ws.Shapes.AddShape(msoShapeDiamond, centerX - (sizeVal / 2), centerY - (sizeVal / 2), sizeVal, sizeVal)
    shp.Name = DASH_PREFIX & "MILESTONE"
    shp.Fill.ForeColor.RGB = colorVal
    shp.Line.Visible = msoFalse

End Sub

'------------------------------------------------------------------------------
' FR: Retourne la valeur Mini Gantt Current Color sans modifier les donnees d'entree.
' EN: Returns the Mini Gantt Current Color value without mutating input data.
'------------------------------------------------------------------------------

Private Function Dashboard_MiniGanttCurrentColor(ByVal progressVal As Double) As Long

    If progressVal >= 1# Then
        Dashboard_MiniGanttCurrentColor = RGB(112, 180, 71)
    Else
        Dashboard_MiniGanttCurrentColor = RGB(68, 114, 196)
    End If

End Function

'------------------------------------------------------------------------------
' FR: Ajoute la reference Progress Bar a la structure cible fournie par l'appelant.
' EN: Adds the Progress Bar reference to the target structure supplied by the caller.
' FR - Effet de bord : cree ou met a jour des shapes Excel.
' EN - Side effect: creates or updates Excel shapes.
'------------------------------------------------------------------------------

Private Sub Dashboard_AddProgressBar(ByVal ws As Worksheet, ByVal x As Double, ByVal y As Double, ByVal w As Double, ByVal h As Double, ByVal progressVal As Double, ByVal startVal As Variant, ByVal finishVal As Variant)

    Dim baseColor As Long
    Dim progressColor As Variant
    Dim progressWidth As Double
    Dim shp As Shape

    baseColor = Dashboard_MiniGanttCurrentColor(progressVal)

    Dashboard_AddBar ws, x, y, w, h, baseColor

    If progressVal > 0# And progressVal < 1# Then
        progressColor = Dashboard_ProgressFillColor(startVal, finishVal, progressVal)
        If Not IsEmpty(progressColor) Then
            progressWidth = w * progressVal
            If progressWidth >= 6 Then
                If progressWidth > w - 2 Then progressWidth = w - 2
                Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, x + 1, y + 1, progressWidth, h - 2)
                shp.Name = DASH_PREFIX & "BAR_PROGRESS"
                shp.Fill.ForeColor.RGB = CLng(progressColor)
                shp.Fill.Transparency = 0.25
                shp.Line.Visible = msoFalse
            End If
        End If
    End If

End Sub

'------------------------------------------------------------------------------
' FR: Retourne la valeur Progress Fill Color sans modifier les donnees d'entree.
' EN: Returns the Progress Fill Color value without mutating input data.
'------------------------------------------------------------------------------

Private Function Dashboard_ProgressFillColor(ByVal startVal As Variant, ByVal finishVal As Variant, ByVal progressVal As Double) As Variant

    Dim durationVal As Double
    Dim coveredDays As Long
    Dim progressFinish As Double

    If progressVal <= 0# Then Exit Function
    If progressVal >= 1# Then
        Dashboard_ProgressFillColor = RGB(112, 180, 71)
        Exit Function
    End If
    If Not Dashboard_HasDateValue(startVal) Or Not Dashboard_HasDateValue(finishVal) Then Exit Function

    durationVal = Dashboard_DateNumber(finishVal) - Dashboard_DateNumber(startVal) + 1
    If durationVal <= 0 Then Exit Function

    coveredDays = Int(durationVal * progressVal + 0.999999)
    If coveredDays < 1 Then coveredDays = 1

    progressFinish = Dashboard_DateNumber(startVal) + coveredDays - 1
    If progressFinish >= CDbl(Date) - 1 Then
        Dashboard_ProgressFillColor = RGB(112, 180, 71)
    Else
        Dashboard_ProgressFillColor = RGB(237, 125, 49)
    End If

End Function

'------------------------------------------------------------------------------
' FR: Actualise Render Time Axis sans modifier les regles metier qui produisent les donnees.
' EN: Refreshes Render Time Axis without changing the business rules that produce the data.
'------------------------------------------------------------------------------

Private Sub Dashboard_RenderTimeAxis(ByVal ws As Worksheet, ByVal minDate As Double, ByVal maxDate As Double, ByVal x As Double, ByVal y As Double, ByVal w As Double)

    Dim spanDays As Double
    Dim tickDate As Date
    Dim tickEnd As Date
    Dim tickX As Double
    Dim stepKind As String
    Dim stepCount As Long
    Dim labelText As String

    If maxDate <= minDate Then Exit Sub

    Dashboard_AddLine ws, x, y + 12, x + w, y + 12, RGB(224, 230, 237), 0.75

    spanDays = maxDate - minDate
    If spanDays <= 14 Then
        stepKind = "d"
        stepCount = 1
        tickDate = CDate(Int(minDate))
    ElseIf spanDays <= 62 Then
        stepKind = "ww"
        stepCount = 1
        tickDate = DateAdd("d", 1 - Weekday(CDate(Int(minDate)), vbMonday), CDate(Int(minDate)))
    ElseIf spanDays <= 730 Then
        stepKind = "m"
        stepCount = 1
        tickDate = DateSerial(Year(CDate(minDate)), Month(CDate(minDate)), 1)
    Else
        stepKind = "yyyy"
        stepCount = 1
        tickDate = DateSerial(Year(CDate(minDate)), 1, 1)
    End If

    tickEnd = CDate(maxDate)
    Do While CDbl(tickDate) <= CDbl(tickEnd)
        tickX = Dashboard_DateToX(CDbl(tickDate), minDate, maxDate, x, w)
        If tickX >= x - 1 And tickX <= x + w + 1 Then
            Dashboard_AddLine ws, tickX, y + 8, tickX, y + 15, RGB(189, 198, 208), 0.75
            Select Case stepKind
                Case "d"
                    labelText = Dashboard_FormatDate(tickDate, False)
                Case "ww"
                    labelText = Dashboard_L("S", "W") & Format$(tickDate, "ww", vbMonday, vbFirstFourDays)
                Case "m"
                    labelText = Dashboard_FormatMonthShort(tickDate)
                Case Else
                    labelText = Format$(tickDate, "yyyy")
            End Select
            Dashboard_AddAxisTickLabel ws, labelText, tickX - 21, y - 5, 42, 12, RGB(96, 111, 128)
        End If

        tickDate = DateAdd(stepKind, stepCount, tickDate)
    Loop

End Sub

'------------------------------------------------------------------------------
' FR: Ajoute la reference Line a la structure cible fournie par l'appelant.
' EN: Adds the Line reference to the target structure supplied by the caller.
' FR - Effet de bord : cree ou met a jour des shapes Excel.
' EN - Side effect: creates or updates Excel shapes.
'------------------------------------------------------------------------------

Private Sub Dashboard_AddLine(ByVal ws As Worksheet, ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double, ByVal colorVal As Long, ByVal weightVal As Double)

    Dim shp As Shape

    Set shp = ws.Shapes.AddLine(x1, y1, x2, y2)
    shp.Name = DASH_PREFIX & "LINE"
    shp.Line.ForeColor.RGB = colorVal
    shp.Line.Weight = weightVal
    shp.Line.DashStyle = msoLineDash

End Sub

'------------------------------------------------------------------------------
' FR: Normalise ou formate Format Text Box selon le contrat canonique du composant.
' EN: Normalizes or formats Format Text Box according to the component contract.
'------------------------------------------------------------------------------

Private Sub Dashboard_FormatTextBox(ByVal shp As Shape, ByVal txt As String, ByVal fontSize As Double, ByVal colorVal As Long, ByVal isBold As Boolean)

    shp.Name = DASH_PREFIX & "TXT"
    shp.Fill.Visible = msoFalse
    shp.Line.Visible = msoFalse
    shp.TextFrame2.MarginLeft = 0
    shp.TextFrame2.MarginRight = 0
    shp.TextFrame2.MarginTop = 0
    shp.TextFrame2.MarginBottom = 0
    shp.TextFrame2.TextRange.Text = txt
    shp.TextFrame2.TextRange.Font.Name = "Segoe UI"
    shp.TextFrame2.TextRange.Font.Size = fontSize
    shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = colorVal
    shp.TextFrame2.TextRange.Font.Bold = IIf(isBold, msoTrue, msoFalse)

End Sub

'------------------------------------------------------------------------------
' FR: Normalise ou formate Format Date Short selon le contrat canonique du composant.
' EN: Normalizes or formats Format Date Short according to the component contract.
'------------------------------------------------------------------------------

Private Function Dashboard_FormatDateShort(ByVal dateVal As Variant) As String

    If Dashboard_HasDateValue(dateVal) Then
        Dashboard_FormatDateShort = Dashboard_FormatDate(dateVal, True)
    Else
        Dashboard_FormatDateShort = "-"
    End If

End Function

'------------------------------------------------------------------------------
' FR: Indique si la valeur Date Value satisfait la condition attendue, sans modifier les donnees source.
' EN: Returns whether the Date Value value satisfies the expected condition without mutating source data.
'------------------------------------------------------------------------------

Private Function Dashboard_HasDateValue(ByVal dateVal As Variant) As Boolean

    Dim v As Variant

    v = GetCellValue(dateVal)
    If Not HasValue(v) Then Exit Function

    Dashboard_HasDateValue = (IsNumeric(v) Or IsDate(v))

End Function

'------------------------------------------------------------------------------
' FR: Retourne la valeur Date Number sans modifier les donnees d'entree.
' EN: Returns the Date Number value without mutating input data.
'------------------------------------------------------------------------------

Private Function Dashboard_DateNumber(ByVal dateVal As Variant) As Double

    If IsDate(dateVal) Then
        Dashboard_DateNumber = CDbl(CDate(dateVal))
    Else
        Dashboard_DateNumber = CDbl(dateVal)
    End If

End Function

'------------------------------------------------------------------------------
' FR: Normalise ou formate Format Signed Days selon le contrat canonique du composant.
' EN: Normalizes or formats Format Signed Days according to the component contract.
'------------------------------------------------------------------------------

Private Function Dashboard_FormatSignedDays(ByVal dayVal As Variant) As String

    If Not HasValue(dayVal) Then
        Dashboard_FormatSignedDays = "-"
    ElseIf CLng(dayVal) > 0 Then
        Dashboard_FormatSignedDays = "+" & CStr(CLng(dayVal)) & Dashboard_L(" jours", " days")
    Else
        Dashboard_FormatSignedDays = CStr(CLng(dayVal)) & Dashboard_L(" jours", " days")
    End If

End Function

'------------------------------------------------------------------------------
' FR: Normalise ou formate Format Signed Compact Days selon le contrat canonique du composant.
' EN: Normalizes or formats Format Signed Compact Days according to the component contract.
'------------------------------------------------------------------------------

Private Function Dashboard_FormatSignedCompactDays(ByVal dayVal As Variant) As String

    If Not HasValue(dayVal) Then
        Dashboard_FormatSignedCompactDays = "-"
    ElseIf CLng(dayVal) > 0 Then
        Dashboard_FormatSignedCompactDays = "+" & CStr(CLng(dayVal)) & Dashboard_L("j", "d")
    Else
        Dashboard_FormatSignedCompactDays = CStr(CLng(dayVal)) & Dashboard_L("j", "d")
    End If

End Function

'------------------------------------------------------------------------------
' FR: Normalise ou formate Format Signed Number selon le contrat canonique du composant.
' EN: Normalizes or formats Format Signed Number according to the component contract.
'------------------------------------------------------------------------------

Private Function Dashboard_FormatSignedNumber(ByVal valueVal As Variant) As String

    If Not HasValue(valueVal) Then
        Dashboard_FormatSignedNumber = "-"
    ElseIf CLng(valueVal) > 0 Then
        Dashboard_FormatSignedNumber = "+" & CStr(CLng(valueVal))
    Else
        Dashboard_FormatSignedNumber = CStr(CLng(valueVal))
    End If

End Function

'------------------------------------------------------------------------------
' FR: Normalise ou formate Format Momentum Delay Percent selon le contrat canonique du composant.
' EN: Normalizes or formats Format Momentum Delay Percent according to the component contract.
'------------------------------------------------------------------------------

Private Function Dashboard_FormatMomentumDelayPercent(ByVal valueVal As Double) As String

    If Abs(valueVal) < 0.01 Then
        If valueVal > 0 Then
            Dashboard_FormatMomentumDelayPercent = "+" & Format$(valueVal, "0.0%")
        Else
            Dashboard_FormatMomentumDelayPercent = Format$(valueVal, "0.0%")
        End If
    Else
        Dashboard_FormatMomentumDelayPercent = Dashboard_FormatPercentSigned(valueVal)
    End If

End Function

'------------------------------------------------------------------------------
' FR: Normalise ou formate Format Percent Signed selon le contrat canonique du composant.
' EN: Normalizes or formats Format Percent Signed according to the component contract.
'------------------------------------------------------------------------------

Private Function Dashboard_FormatPercentSigned(ByVal valueVal As Double) As String

    If valueVal > 0 Then
        Dashboard_FormatPercentSigned = "+" & Format$(valueVal, "0%")
    Else
        Dashboard_FormatPercentSigned = Format$(valueVal, "0%")
    End If

End Function

'------------------------------------------------------------------------------
' FR: Retourne la valeur Language sans exposer de mutateur sur l'etat source.
' EN: Returns the Language value without exposing a mutator for source state.
'------------------------------------------------------------------------------

Public Function Dashboard_CurrentLanguage() As String

    If UCase$(Trim$(gDashboardLanguage)) <> "FR" And UCase$(Trim$(gDashboardLanguage)) <> "EN" Then
        gDashboardLanguage = "EN"
    End If
    Dashboard_CurrentLanguage = UCase$(Trim$(gDashboardLanguage))

End Function

'------------------------------------------------------------------------------
' FR: Indique si la valeur French satisfait la condition attendue, sans modifier les donnees source.
' EN: Returns whether the French value satisfies the expected condition without mutating source data.
'------------------------------------------------------------------------------

Private Function Dashboard_IsFrench() As Boolean

    Dashboard_IsFrench = (Dashboard_CurrentLanguage() = "FR")

End Function

'------------------------------------------------------------------------------
' FR: Retourne la valeur L sans modifier les donnees d'entree.
' EN: Returns the L value without mutating input data.
'------------------------------------------------------------------------------

Private Function Dashboard_L(ByVal frText As String, ByVal enText As String) As String

    If Dashboard_IsFrench() Then
        Dashboard_L = frText
    Else
        Dashboard_L = enText
    End If

End Function

'------------------------------------------------------------------------------
' FR: Retourne la valeur No Project Loaded Text sans modifier les donnees d'entree.
' EN: Returns the No Project Loaded Text value without mutating input data.
'------------------------------------------------------------------------------

Private Function Dashboard_NoProjectLoadedText() As String

    Dashboard_NoProjectLoadedText = Dashboard_L("Aucun projet chargé", "No project loaded")

End Function

'------------------------------------------------------------------------------
' FR: Retourne la valeur No Data Text sans modifier les donnees d'entree.
' EN: Returns the No Data Text value without mutating input data.
'------------------------------------------------------------------------------

Private Function Dashboard_NoDataText() As String

    Dashboard_NoDataText = Dashboard_L("AUCUNE DONNÉE", "NO DATA")

End Function

'------------------------------------------------------------------------------
' FR: Retourne la valeur Duration Suffix sans modifier les donnees d'entree.
' EN: Returns the Duration Suffix value without mutating input data.
'------------------------------------------------------------------------------

Private Function Dashboard_DurationSuffix() As String

    Dashboard_DurationSuffix = Dashboard_L("j", "d")

End Function

'------------------------------------------------------------------------------
' FR: Normalise ou formate Format Month Short selon le contrat canonique du composant.
' EN: Normalizes or formats Format Month Short according to the component contract.
'------------------------------------------------------------------------------

Private Function Dashboard_FormatMonthShort(ByVal dateVal As Variant) As String

    Dim monthNames As Variant
    Dim d As Date

    If Not Dashboard_HasDateValue(dateVal) Then Exit Function
    d = CDate(Dashboard_DateNumber(dateVal))

    If Dashboard_IsFrench() Then
        monthNames = Array("janv", "fév", "mars", "avr", "mai", "juin", "juil", "août", "sept", "oct", "nov", "déc")
    Else
        monthNames = Array("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")
    End If

    Dashboard_FormatMonthShort = CStr(monthNames(Month(d) - 1))

End Function

'------------------------------------------------------------------------------
' FR: Normalise ou formate Format Date selon le contrat canonique du composant.
' EN: Normalizes or formats Format Date according to the component contract.
'------------------------------------------------------------------------------

Private Function Dashboard_FormatDate(ByVal dateVal As Variant, Optional ByVal includeYear As Boolean = True) As String

    Dim d As Date

    If Not Dashboard_HasDateValue(dateVal) Then
        Dashboard_FormatDate = "-"
        Exit Function
    End If

    d = CDate(Dashboard_DateNumber(dateVal))
    Dashboard_FormatDate = Format$(d, "dd") & " " & Dashboard_FormatMonthShort(d)
    If includeYear Then Dashboard_FormatDate = Dashboard_FormatDate & " " & Format$(d, "yyyy")

End Function

'------------------------------------------------------------------------------
' FR: Normalise ou formate Format Timestamp selon le contrat canonique du composant.
' EN: Normalizes or formats Format Timestamp according to the component contract.
'------------------------------------------------------------------------------

Private Function Dashboard_FormatTimestamp(ByVal dateVal As Date) As String

    Dashboard_FormatTimestamp = Dashboard_L("Mis à jour ", "Updated ") & Dashboard_FormatDate(dateVal, True) & " " & Format$(dateVal, "hh:nn")

End Function

'------------------------------------------------------------------------------
' FR: Indique si la valeur Timestamp Text satisfait la condition attendue, sans modifier les donnees source.
' EN: Returns whether the Timestamp Text value satisfies the expected condition without mutating source data.
'------------------------------------------------------------------------------

Private Function Dashboard_IsTimestampText(ByVal txt As String) As Boolean

    Dashboard_IsTimestampText = (Left$(txt, 8) = "Updated " Or Left$(txt, 11) = "Mis à jour ")

End Function

'------------------------------------------------------------------------------
' FR: Retourne la valeur Chart Date Number Format sans modifier les donnees d'entree.
' EN: Returns the Chart Date Number Format value without mutating input data.
'------------------------------------------------------------------------------

Private Function Dashboard_ChartDateNumberFormat(Optional ByVal includeYear As Boolean = False) As String

    If includeYear Then
        Dashboard_ChartDateNumberFormat = Settings_GetDateNumberFormat()
    Else
        Dashboard_ChartDateNumberFormat = Settings_GetCompactDateNumberFormat()
    End If

End Function

'------------------------------------------------------------------------------
' FR: Retourne la valeur Progress Color sans modifier les donnees d'entree.
' EN: Returns the Progress Color value without mutating input data.
'------------------------------------------------------------------------------

Private Function Dashboard_ProgressColor(ByVal progressVar As Double) As Long

    If progressVar < -0.03 Then
        Dashboard_ProgressColor = RGB(192, 80, 77)
    ElseIf progressVar < 0 Then
        Dashboard_ProgressColor = RGB(238, 156, 68)
    Else
        Dashboard_ProgressColor = RGB(0, 145, 112)
    End If

End Function

'------------------------------------------------------------------------------
' FR: Retourne la valeur Drift Color sans modifier les donnees d'entree.
' EN: Returns the Drift Color value without mutating input data.
'------------------------------------------------------------------------------

Private Function Dashboard_DriftColor(ByVal driftDays As Variant) As Long

    If Not HasValue(driftDays) Then
        Dashboard_DriftColor = RGB(96, 111, 128)
    ElseIf CLng(driftDays) > 0 Then
        Dashboard_DriftColor = RGB(192, 80, 77)
    Else
        Dashboard_DriftColor = RGB(0, 145, 112)
    End If

End Function
