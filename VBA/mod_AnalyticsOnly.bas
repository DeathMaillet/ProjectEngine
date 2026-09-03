Attribute VB_Name = "mod_AnalyticsOnly"
Option Explicit

'===============================================================================
' MODULE : mod_AnalyticsOnly
' DOMAINE / DOMAIN : Analytics
'
' FR
' Pilote l'activation des analytics, leur execution isolee et le nettoyage de leurs sorties.
' Ne doit pas contourner les contrats publics des autres domaines.
'
' EN
' Controls analytics enablement, standalone execution and output cleanup.
' Must not bypass public contracts owned by other domains.
'
' CONTRATS / CONTRACTS : IsAnalyticsEnabled, SetAnalyticsEnabled, Toggle_Analytics, Ensure_Analytics_Toggle, Refresh_Analytics_Toggle_Visual, Run_Analytics_Only, Clear_Analytics_Outputs
' CALLBACKS EXTERNES / EXTERNAL CALLBACKS : Aucun / None
'===============================================================================


' ============================================================
' mod_AnalyticsOnly
' ANALYTICS TOGGLE + ANALYTICS ONLY EXECUTION
'
' Objectif :
' - Le toggle visible est dans WBS, au-dessus des colonnes analytics.
' - L'état ON/OFF est conservé en mémoire VBA uniquement.
' - CALC!AB1 n'est plus utilisé comme backend state.
'
' Comportement :
'
' OFF :
' - ne recalcule pas Critical Path / floats / REX ;
' - vide les colonnes analytics dans WBS ;
' - vide aussi les colonnes analytics dans CALC si elles existent.
'
' ON :
' - recalcule uniquement les analytics sur l'état courant de CALC ;
' - pousse les analytics vers WBS ;
' - ne relance pas le core planning complet.
'
' Principe important :
' - Pas de duplication du moteur analytics.
' - CalcBridge_RunAnalyticsAndPush reste la seule source
'   de calcul CP / floats / REX.
'
' Shapes UI dans WBS :
' - lbl_WBS_Analytics_Toggle
' - shp_WBS_Analytics_BG
' - shp_WBS_Analytics_Knob
'
' Colonnes analytics concernées :
' - Critical Path
' - Critical Path REX
' - Total Float
' - Free Float
' - Total Float REX
' - Free Float REX
' - Deadline Float
' ============================================================

Private gAnalyticsEnabled As Boolean
Private gAnalyticsStateInitialized As Boolean

Private Const ANALYTICS_TOGGLE_WS As String = "WBS"
Private Const ANALYTICS_LABEL_CELL As String = "AB1"
Private Const ANALYTICS_TOGGLE_CELL As String = "AC1"

Private Const ANALYTICS_LABEL_SHAPE As String = "lbl_WBS_Analytics_Toggle"
Private Const ANALYTICS_BG_SHAPE As String = "shp_WBS_Analytics_BG"
Private Const ANALYTICS_KNOB_SHAPE As String = "shp_WBS_Analytics_Knob"

'------------------------------------------------------------------------------
' FR: Indique si Analytics Enabled est vrai pour le contexte courant.
' EN: Returns whether Analytics Enabled is true for the current context.
'------------------------------------------------------------------------------
Public Function IsAnalyticsEnabled() As Boolean

    If Not gAnalyticsStateInitialized Then
        gAnalyticsEnabled = True
        gAnalyticsStateInitialized = True
    End If

    IsAnalyticsEnabled = gAnalyticsEnabled

End Function

'------------------------------------------------------------------------------
' FR: Met a jour Analytics Enabled dans le contexte analytics.
' EN: Updates Analytics Enabled in the analytics context.
'------------------------------------------------------------------------------
Public Sub SetAnalyticsEnabled(ByVal enabledValue As Boolean)

    gAnalyticsEnabled = enabledValue
    gAnalyticsStateInitialized = True

End Sub

'------------------------------------------------------------------------------
' FR: Bascule l'etat Analytics et met a jour les sorties associees.
' EN: Toggles Analytics state and updates related outputs.
'------------------------------------------------------------------------------
Public Sub Toggle_Analytics()

    Dim enabledNow As Boolean
    Dim enabledNext As Boolean
    Dim consoleMessages As Collection

    On Error GoTo ErrHandler

    Set consoleMessages = New Collection

    enabledNow = IsAnalyticsEnabled()
    enabledNext = Not enabledNow

    SetAnalyticsEnabled enabledNext
    Refresh_Analytics_Toggle_Visual

    If enabledNext Then
        Run_Analytics_Only consoleMessages
    Else
        Clear_Analytics_Outputs consoleMessages
    End If

    Refresh_Analytics_Toggle_Visual
    CalcBridge_ShowPlanningConsole consoleMessages

    Exit Sub

ErrHandler:
    CalcBridge_ShowSingleConsoleMessage _
        "STOP", _
        "Erreur dans Toggle_Analytics : " & Err.Description, _
        "Error in Toggle_Analytics: " & Err.Description

End Sub

'------------------------------------------------------------------------------
' FR: Verifie ou cree Analytics Toggle si necessaire.
' EN: Ensures or creates Analytics Toggle when needed.
'------------------------------------------------------------------------------
Public Sub Ensure_Analytics_Toggle()

    On Error GoTo SafeExit

    Ensure_Analytics_Toggle_Shapes
    Refresh_Analytics_Toggle_Visual

SafeExit:
End Sub

'------------------------------------------------------------------------------
' FR: Rafraichit Analytics Toggle Visual a partir de l'etat courant.
' EN: Refreshes Analytics Toggle Visual from the current state.
'------------------------------------------------------------------------------
Public Sub Refresh_Analytics_Toggle_Visual()

    Dim ws As Worksheet
    Dim shpLabel As Shape
    Dim shpBg As Shape
    Dim shpKnob As Shape
    Dim isOn As Boolean
    Dim knobLeft As Double

    On Error GoTo SafeExit

    Set ws = ThisWorkbook.Worksheets(ANALYTICS_TOGGLE_WS)

    On Error Resume Next
    Set shpLabel = ws.Shapes(ANALYTICS_LABEL_SHAPE)
    Set shpBg = ws.Shapes(ANALYTICS_BG_SHAPE)
    Set shpKnob = ws.Shapes(ANALYTICS_KNOB_SHAPE)
    On Error GoTo SafeExit

    If shpBg Is Nothing Or shpKnob Is Nothing Then
        Ensure_Analytics_Toggle_Shapes

        On Error Resume Next
        Set shpLabel = ws.Shapes(ANALYTICS_LABEL_SHAPE)
        Set shpBg = ws.Shapes(ANALYTICS_BG_SHAPE)
        Set shpKnob = ws.Shapes(ANALYTICS_KNOB_SHAPE)
        On Error GoTo SafeExit
    End If

    If shpBg Is Nothing Then Exit Sub
    If shpKnob Is Nothing Then Exit Sub

    isOn = IsAnalyticsEnabled()

    If Not shpLabel Is Nothing Then
        With shpLabel
            .TextFrame2.TextRange.Text = "Analytics"
            .TextFrame2.TextRange.Font.Size = 9.5
            .TextFrame2.TextRange.Font.Bold = msoTrue
            .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
            .TextFrame2.VerticalAnchor = msoAnchorMiddle
            .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignLeft
            .OnAction = "Toggle_Analytics"
        End With
    End If

    If isOn Then
        shpBg.Fill.ForeColor.RGB = RGB(112, 173, 71)
        shpBg.Line.ForeColor.RGB = RGB(112, 173, 71)
        knobLeft = shpBg.Left + shpBg.Width - shpKnob.Width - 3
    Else
        shpBg.Fill.ForeColor.RGB = RGB(230, 230, 230)
        shpBg.Line.ForeColor.RGB = RGB(170, 170, 170)
        knobLeft = shpBg.Left + 3
    End If

    shpBg.OnAction = "Toggle_Analytics"
    shpKnob.OnAction = "Toggle_Analytics"

    shpKnob.Left = knobLeft
    shpKnob.Top = shpBg.Top + ((shpBg.Height - shpKnob.Height) / 2)

SafeExit:
End Sub


'------------------------------------------------------------------------------
' FR: Lance le workflow Analytics Only.
' EN: Runs the Analytics Only workflow.
'------------------------------------------------------------------------------
Public Sub Run_Analytics_Only( _
    Optional ByVal consoleMessages As Collection = Nothing, _
    Optional ByVal executionNetwork As clsCompiledExecutionNetwork = Nothing, _
    Optional ByVal sharedLinksBySuccId As Object = Nothing)

    Dim perfScope As clsPerfScope

    Dim wsCalc As Worksheet
    Dim tblCalc As ListObject
    Dim mapCalc As Object
    Dim linksBySuccId As Object
    Dim localConsole As Collection
    Dim showLocalConsole As Boolean

    Set perfScope = Profiler_BeginScope("Run_Analytics_Only", "Analytics")

    On Error GoTo ErrHandler

    If consoleMessages Is Nothing Then
        Set localConsole = New Collection
        showLocalConsole = True
    Else
        Set localConsole = consoleMessages
        showLocalConsole = False
    End If

    Set wsCalc = ThisWorkbook.Worksheets("CALC")
    Set tblCalc = wsCalc.ListObjects("tbl_CALC")

    If tblCalc.DataBodyRange Is Nothing Then
        CalcBridge_AddConsoleMessage localConsole, "WARNING", _
            BiMsg( _
                "Analytics non recalculées : tbl_CALC est vide.", _
                "Analytics not recalculated: tbl_CALC is empty.")

        If showLocalConsole Then CalcBridge_ShowPlanningConsole localConsole
        Exit Sub
    End If

    Set mapCalc = CanonicalIdentity_BuildColumnMap(tblCalc)
    If sharedLinksBySuccId Is Nothing Then
        Set linksBySuccId = BuildCoreLinksBySucc_FromLogicLinksTable_Expanded(tblCalc)
    Else
        Set linksBySuccId = sharedLinksBySuccId
    End If

    CalcBridge_RunAnalyticsAndPush tblCalc, mapCalc, linksBySuccId, localConsole, executionNetwork
    If AnalyticsOnly_TableHasNegativeFloat(tblCalc, mapCalc) Then
        CalcBridge_AddConsoleMessage localConsole, "WARNING", AnalyticsOnly_NegativeFloatWarningMessage()
    End If
    CalcBridge_ComputeDeadlineAnalytics tblCalc, mapCalc, localConsole

    Push_Analytics_Back_To_WBS

    If showLocalConsole Then
        CalcBridge_ShowPlanningConsole localConsole
    End If

    Exit Sub

ErrHandler:
    If localConsole Is Nothing Then Set localConsole = New Collection

    CalcBridge_AddConsoleMessage localConsole, "STOP", _
        BiMsg( _
            "Erreur dans Run_Analytics_Only" & vbCrLf & _
            "-> " & Err.Description, _
            "Error in Run_Analytics_Only" & vbCrLf & _
            "-> " & Err.Description)

    CalcBridge_ShowPlanningConsole localConsole

End Sub

'------------------------------------------------------------------------------
' FR: Vide ou reinitialise Analytics Outputs.
' EN: Clears or resets Analytics Outputs.
'------------------------------------------------------------------------------
Public Sub Clear_Analytics_Outputs(Optional ByVal consoleMessages As Collection = Nothing)

    Dim wsCalc As Worksheet
    Dim wsWBS As Worksheet
    Dim tblCalc As ListObject
    Dim tblWBS As ListObject
    Dim analyticsCols As Variant
    Dim i As Long
    Dim allowedFields As Variant
    Dim localConsole As Collection
    Dim showLocalConsole As Boolean
    Dim writeScopeToken As Long
    Dim errorDescription As String

    On Error GoTo ErrHandler

    If consoleMessages Is Nothing Then
        Set localConsole = New Collection
        showLocalConsole = True
    Else
        Set localConsole = consoleMessages
        showLocalConsole = False
    End If

    analyticsCols = AnalyticsColumnNames()

    Set wsCalc = ThisWorkbook.Worksheets("CALC")
    Set wsWBS = ThisWorkbook.Worksheets("WBS")

    Set tblCalc = wsCalc.ListObjects("tbl_CALC")
    Set tblWBS = wsWBS.ListObjects("tbl_WBS")

    If Not tblCalc.DataBodyRange Is Nothing Then
        For i = LBound(analyticsCols) To UBound(analyticsCols)
            ClearListColumnIfExists tblCalc, _
                SchemaCanonicalEnglishColumnTitle(VTS_TABLE_WBS, CStr(analyticsCols(i)))
        Next i
    End If

    allowedFields = AnalyticsVisibleColumnTitles(analyticsCols)

    If Not tblWBS.DataBodyRange Is Nothing Then

        writeScopeToken = OpenAuthorizedWBSWriteScope( _
            "Clear_Analytics_Outputs", allowedFields)

        For i = LBound(analyticsCols) To UBound(analyticsCols)
            ClearVisibleWBSColumn tblWBS, CStr(analyticsCols(i))
        Next i

        CloseAuthorizedWBSWriteScope writeScopeToken
        writeScopeToken = 0

    End If

    Exit Sub

ErrHandler:
    errorDescription = Err.Description
    On Error Resume Next
    CloseAuthorizedWBSWriteScope writeScopeToken
    On Error GoTo 0

    If localConsole Is Nothing Then Set localConsole = New Collection

    CalcBridge_AddConsoleMessage localConsole, "STOP", _
        BiMsg( _
            "Erreur dans Clear_Analytics_Outputs" & vbCrLf & _
            "-> " & errorDescription, _
            "Error in Clear_Analytics_Outputs" & vbCrLf & _
            "-> " & errorDescription)

    CalcBridge_ShowPlanningConsole localConsole

End Sub


'------------------------------------------------------------------------------
' FR: Retourne la map Analytics Only Table Has Negative Float sans modifier les donnees d'entree.
' EN: Returns the Analytics Only Table Has Negative Float map without mutating input data.
'------------------------------------------------------------------------------

Private Function AnalyticsOnly_TableHasNegativeFloat( _
    ByVal tblCalc As ListObject, _
    ByVal mapCalc As Object) As Boolean

    Dim columnsToCheck As Variant
    Dim colName As Variant
    Dim arr As Variant
    Dim r As Long

    If tblCalc Is Nothing Then Exit Function
    If tblCalc.DataBodyRange Is Nothing Then Exit Function
    If mapCalc Is Nothing Then Exit Function

    columnsToCheck = Array("Total Float", "Free Float", "Total Float REX", "Free Float REX")
    arr = tblCalc.DataBodyRange.value

    For Each colName In columnsToCheck
        If mapCalc.Exists(CStr(colName)) Then
            For r = 1 To UBound(arr, 1)
                If HasValue(arr(r, mapCalc(CStr(colName)))) Then
                    If IsNumeric(arr(r, mapCalc(CStr(colName)))) Then
                        If CDbl(arr(r, mapCalc(CStr(colName)))) < 0 Then
                            AnalyticsOnly_TableHasNegativeFloat = True
                            Exit Function
                        End If
                    End If
                End If
            Next r
        End If
    Next colName

End Function

'------------------------------------------------------------------------------
' FR: Retourne la valeur Analytics Only Negative Float Warning Message sans modifier les donnees d'entree.
' EN: Returns the Analytics Only Negative Float Warning Message value without mutating input data.
'------------------------------------------------------------------------------

Private Function AnalyticsOnly_NegativeFloatWarningMessage() As String

    AnalyticsOnly_NegativeFloatWarningMessage = BiMsg( _
        "Float négatif détecté dans le planning actuel" & vbCrLf & _
        "-> vérifier la logique, les dates, les lags ou les prévisions", _
        "Negative float detected in the current schedule" & vbCrLf & _
        "-> check logic, dates, lags or forecasts")

End Function

'------------------------------------------------------------------------------
' FR: Retourne la valeur Analytics Column Names sans modifier les donnees d'entree.
' EN: Returns the Analytics Column Names value without mutating input data.
'------------------------------------------------------------------------------

Private Function AnalyticsColumnNames() As Variant

    AnalyticsColumnNames = Array( _
        VTS_COL_CRITICAL_PATH, _
        VTS_COL_LONGEST_PATH, _
        VTS_COL_CRITICAL_PATH_REX, _
        VTS_COL_TOTAL_FLOAT, _
        VTS_COL_FREE_FLOAT, _
        VTS_COL_TOTAL_FLOAT_REX, _
        VTS_COL_FREE_FLOAT_REX, _
        VTS_COL_DEADLINE_FLOAT _
    )

End Function

Private Function AnalyticsVisibleColumnTitles(ByVal columnKeys As Variant) As Variant
    Dim titles() As Variant
    Dim i As Long

    ReDim titles(LBound(columnKeys) To UBound(columnKeys))
    For i = LBound(columnKeys) To UBound(columnKeys)
        titles(i) = SchemaCurrentColumnTitle(VTS_TABLE_WBS, CStr(columnKeys(i)))
    Next i

    AnalyticsVisibleColumnTitles = titles
End Function

Private Sub ClearVisibleWBSColumn(ByVal tblWBS As ListObject, ByVal columnKey As String)
    Dim col As ListColumn

    Set col = SchemaListColumn(tblWBS, VTS_TABLE_WBS, columnKey)
    If Not col.DataBodyRange Is Nothing Then col.DataBodyRange.ClearContents
End Sub

'------------------------------------------------------------------------------
' FR: Vide ou reinitialise List Column If Exists.
' EN: Clears or resets List Column If Exists.
'------------------------------------------------------------------------------
Private Sub ClearListColumnIfExists(ByVal tbl As ListObject, ByVal columnName As String)

    Dim col As ListColumn

    On Error Resume Next
    Set col = tbl.ListColumns(columnName)
    On Error GoTo 0

    If Not col Is Nothing Then
        If Not col.DataBodyRange Is Nothing Then
            col.DataBodyRange.ClearContents
        End If
    End If

End Sub

'------------------------------------------------------------------------------
' FR: Verifie ou cree Analytics Toggle Shapes si necessaire.
' EN: Ensures or creates Analytics Toggle Shapes when needed.
'------------------------------------------------------------------------------
Private Sub Ensure_Analytics_Toggle_Shapes()

    Dim ws As Worksheet
    Dim labelCell As Range
    Dim toggleCell As Range
    Dim shpLabel As Shape
    Dim shpBg As Shape
    Dim shpKnob As Shape

    Dim labelLeft As Double
    Dim labelTop As Double
    Dim labelWidth As Double
    Dim labelHeight As Double

    Dim bgLeft As Double
    Dim bgTop As Double
    Dim bgWidth As Double
    Dim bgHeight As Double
    Dim knobSize As Double

    On Error GoTo SafeExit

    Set ws = ThisWorkbook.Worksheets(ANALYTICS_TOGGLE_WS)
    Set labelCell = ws.Range(ANALYTICS_LABEL_CELL)
    Set toggleCell = ws.Range(ANALYTICS_TOGGLE_CELL)

    With labelCell
        .ClearContents
        .Interior.Pattern = xlNone
        .Borders.LineStyle = xlNone
    End With

    With toggleCell
        .ClearContents
        .Interior.Pattern = xlNone
        .Borders.LineStyle = xlNone
    End With

    labelLeft = labelCell.Left + 4
    labelTop = labelCell.Top + 2
    labelWidth = labelCell.Width - 4
    labelHeight = labelCell.Height - 4

    bgWidth = 42
    bgHeight = 20
    knobSize = 14

    bgLeft = toggleCell.Left + ((toggleCell.Width - bgWidth) / 2)
    bgTop = toggleCell.Top + ((toggleCell.Height - bgHeight) / 2)

    On Error Resume Next
    Set shpLabel = ws.Shapes(ANALYTICS_LABEL_SHAPE)
    Set shpBg = ws.Shapes(ANALYTICS_BG_SHAPE)
    Set shpKnob = ws.Shapes(ANALYTICS_KNOB_SHAPE)
    On Error GoTo SafeExit

    If shpLabel Is Nothing Then
        Set shpLabel = ws.Shapes.AddTextbox( _
            msoTextOrientationHorizontal, _
            labelLeft, _
            labelTop, _
            labelWidth, _
            labelHeight)
        shpLabel.Name = ANALYTICS_LABEL_SHAPE
    End If

    With shpLabel
        .Left = labelLeft
        .Top = labelTop
        .Width = labelWidth
        .Height = labelHeight
        .Line.Visible = msoFalse
        .Fill.Visible = msoFalse
        .Placement = xlMoveAndSize
        .OnAction = "Toggle_Analytics"
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .TextFrame2.MarginLeft = 0
        .TextFrame2.MarginRight = 0
        .TextFrame2.MarginTop = 0
        .TextFrame2.MarginBottom = 0
        .TextFrame2.TextRange.Text = "Analytics"
        .TextFrame2.TextRange.Font.Size = 9.5
        .TextFrame2.TextRange.Font.Bold = msoTrue
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignLeft
    End With

    If shpBg Is Nothing Then
        Set shpBg = ws.Shapes.AddShape( _
            msoShapeRoundedRectangle, _
            bgLeft, _
            bgTop, _
            bgWidth, _
            bgHeight)
        shpBg.Name = ANALYTICS_BG_SHAPE
    End If

    With shpBg
        .Left = bgLeft
        .Top = bgTop
        .Width = bgWidth
        .Height = bgHeight
        .Adjustments.item(1) = 0.5
        .Placement = xlMoveAndSize
        .OnAction = "Toggle_Analytics"
        .Line.Weight = 1
        .Shadow.Visible = msoFalse
    End With

    If shpKnob Is Nothing Then
        Set shpKnob = ws.Shapes.AddShape( _
            msoShapeOval, _
            bgLeft + 3, _
            bgTop + ((bgHeight - knobSize) / 2), _
            knobSize, _
            knobSize)
        shpKnob.Name = ANALYTICS_KNOB_SHAPE
    End If

    With shpKnob
        .Width = knobSize
        .Height = knobSize
        .Placement = xlMoveAndSize
        .OnAction = "Toggle_Analytics"
        .Fill.ForeColor.RGB = RGB(255, 255, 255)
        .Line.ForeColor.RGB = RGB(150, 150, 150)
        .Line.Weight = 0.75
        .Shadow.Visible = msoTrue
        .Shadow.Blur = 4
        .Shadow.OffsetX = 0
        .Shadow.OffsetY = 1
        .Shadow.Transparency = 0.45
    End With

SafeExit:
End Sub

