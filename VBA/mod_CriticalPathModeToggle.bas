Attribute VB_Name = "mod_CriticalPathModeToggle"
Option Explicit

'===============================================================================
' MODULE : mod_CriticalPathModeToggle
' DOMAINE / DOMAIN : Shared Infrastructure
'
' FR
' Possede le toggle Critical/Longest Path partage par Settings et Gantt.
' Ne doit pas contourner les contrats publics des autres domaines.
'
' EN
' Owns the Critical/Longest Path toggle shared by Settings and Gantt.
' Must not bypass public contracts owned by other domains.
'
' CONTRATS / CONTRACTS : IsCriticalPathMultiNetworkEnabled, Toggle_CriticalPathMode, Toggle_CriticalPathMode_FromGantt, Ensure_CriticalPathMode_Toggle, SetCriticalPathModeVisualState, Refresh_CriticalPathMode_Toggle_Visual, Cleanup_CriticalPathMode_Gantt_Legacy_Shapes
' CALLBACKS EXTERNES / EXTERNAL CALLBACKS : Aucun / None
'===============================================================================


Private Const CP_MODE_WS As String = "WBS"
Private Const CP_MODE_LABEL_CELL As String = "AD1"
Private Const CP_MODE_STATE_CELL As String = "AE1"


Private Const CP_MODE_GLOBAL As String = "GLOBAL"
Private Const CP_MODE_MULTI As String = "MULTI NETWORK"

Private Const CP_LABEL_SHAPE As String = "lbl_WBS_CPMode_Toggle"
Private Const CP_BG_SHAPE As String = "shp_WBS_CPMode_BG"
Private Const CP_KNOB_SHAPE As String = "shp_WBS_CPMode_Knob"

Private Const CP_GANTT_WS As String = "GANTT"

Private Const CP_GANTT_LABEL_SHAPE As String = "lbl_GANTT_CPMode_Toggle"
Private Const CP_GANTT_BG_SHAPE As String = "shp_GANTT_CPMode_BG"
Private Const CP_GANTT_KNOB_SHAPE As String = "shp_GANTT_CPMode_Knob"

'------------------------------------------------------------------------------
' FR: Indique si Critical Path Multi Network Enabled est vrai pour le contexte courant.
' EN: Returns whether Critical Path Multi Network Enabled is true for the current context.
'------------------------------------------------------------------------------
Public Function IsCriticalPathMultiNetworkEnabled() As Boolean

    Dim ws As Worksheet
    Dim v As String

    On Error GoTo SafeExit

    Set ws = ThisWorkbook.Worksheets(CP_MODE_WS)

    v = UCase$(Trim$(CStr(ws.Range(CP_MODE_STATE_CELL).value)))
    v = Replace$(v, "-", " ")
    v = Replace$(v, "_", " ")
    v = Replace$(v, "É", "E")

    Do While InStr(1, v, "  ", vbBinaryCompare) > 0
        v = Replace$(v, "  ", " ")
    Loop

    IsCriticalPathMultiNetworkEnabled = _
        (v = CP_MODE_MULTI Or _
         v = "MULTI" Or _
         v = "MULTINETWORK" Or _
         v = "MULTI NETWORKS" Or _
         v = "MULTI RESEAU" Or _
         v = "MULTI RESEAUX" Or _
         v = "ON" Or _
         v = "TRUE" Or _
         v = "1" Or _
         v = "YES" Or _
         v = "OUI")

    Exit Function

SafeExit:
    'Fail-safe volontaire : comportement historique.
    IsCriticalPathMultiNetworkEnabled = False

End Function

'------------------------------------------------------------------------------
' FR: Bascule l'etat Critical Path Mode et met a jour les sorties associees.
' EN: Toggles Critical Path Mode state and updates related outputs.
'------------------------------------------------------------------------------
Public Sub Toggle_CriticalPathMode()

    Toggle_CriticalPathMode_Core False

End Sub

'------------------------------------------------------------------------------
' FR: Bascule l'etat Critical Path Mode From Gantt et met a jour les sorties associees.
' EN: Toggles Critical Path Mode From Gantt state and updates related outputs.
'------------------------------------------------------------------------------
Public Sub Toggle_CriticalPathMode_FromGantt()

    Toggle_CriticalPathMode_Core True

End Sub

'------------------------------------------------------------------------------
' FR: Verifie ou cree Critical Path Mode Toggle si necessaire.
' EN: Ensures or creates Critical Path Mode Toggle when needed.
'------------------------------------------------------------------------------
Public Sub Ensure_CriticalPathMode_Toggle()

    Dim ws As Worksheet
    Dim stateText As String

    On Error GoTo SafeExit

    Set ws = ThisWorkbook.Worksheets(CP_MODE_WS)

    stateText = UCase$(Trim$(CStr(ws.Range(CP_MODE_STATE_CELL).value)))

    If stateText = "" Then
        SetCriticalPathModeVisualState False
    Else
        SetCriticalPathModeVisualState IsCriticalPathMultiNetworkEnabled()
    End If

    Ensure_CriticalPathMode_Shapes
    Refresh_CriticalPathMode_Toggle_Visual

SafeExit:
End Sub

'------------------------------------------------------------------------------
' FR: Met a jour Critical Path Mode Visual State dans le contexte critical path mode.
' EN: Updates Critical Path Mode Visual State in the critical path mode context.
'------------------------------------------------------------------------------
Public Sub SetCriticalPathModeVisualState(ByVal multiNetworkEnabled As Boolean)

    Dim ws As Worksheet

    On Error GoTo SafeExit

    Set ws = ThisWorkbook.Worksheets(CP_MODE_WS)

    'AD/AE restent le backend/fallback.
    'Le texte cellule est masqué pour éviter les doublons derrière les shapes.
    With ws.Range(CP_MODE_LABEL_CELL)
        .value = "Multi Critical Path"
        .NumberFormat = "@"
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Bold = True
        .Interior.Color = RGB(255, 255, 255)
        .Font.Color = .Interior.Color
        .Borders.LineStyle = xlNone
    End With

    With ws.Range(CP_MODE_STATE_CELL)
        .NumberFormat = "@"
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Bold = True
        .Interior.Color = RGB(255, 255, 255)
        .Font.Color = .Interior.Color
        .Borders.LineStyle = xlNone

        If multiNetworkEnabled Then
            .value = CP_MODE_MULTI
        Else
            .value = CP_MODE_GLOBAL
        End If
    End With

SafeExit:
End Sub

'------------------------------------------------------------------------------
' FR: Rafraichit Critical Path Mode Toggle Visual a partir de l'etat courant.
' EN: Refreshes Critical Path Mode Toggle Visual from the current state.
'------------------------------------------------------------------------------
Public Sub Refresh_CriticalPathMode_Toggle_Visual()

    Dim ws As Worksheet
    Dim shpLabel As Shape
    Dim shpBg As Shape
    Dim shpKnob As Shape
    Dim isMulti As Boolean
    Dim knobLeft As Double

    On Error GoTo SafeExit

    Set ws = ThisWorkbook.Worksheets(CP_MODE_WS)

    On Error Resume Next
    Set shpLabel = ws.Shapes(CP_LABEL_SHAPE)
    Set shpBg = ws.Shapes(CP_BG_SHAPE)
    Set shpKnob = ws.Shapes(CP_KNOB_SHAPE)
    On Error GoTo SafeExit

    If shpBg Is Nothing Then Exit Sub
    If shpKnob Is Nothing Then Exit Sub

    isMulti = IsCriticalPathMultiNetworkEnabled()

    If Not shpLabel Is Nothing Then
        With shpLabel
            .TextFrame2.TextRange.Text = "Multi Critical Path"
            .TextFrame2.TextRange.Font.Size = 9.5
            .TextFrame2.TextRange.Font.Bold = msoTrue
            .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
            .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignLeft
            .TextFrame2.VerticalAnchor = msoAnchorMiddle
            .OnAction = ""
        End With
    End If

    If isMulti Then
        shpBg.Fill.ForeColor.RGB = RGB(112, 173, 71)
        shpBg.Line.ForeColor.RGB = RGB(112, 173, 71)
        knobLeft = shpBg.Left + shpBg.Width - shpKnob.Width - 3
    Else
        shpBg.Fill.ForeColor.RGB = RGB(230, 230, 230)
        shpBg.Line.ForeColor.RGB = RGB(170, 170, 170)
        knobLeft = shpBg.Left + 3
    End If

    shpBg.OnAction = "Toggle_CriticalPathMode"
    shpKnob.OnAction = "Toggle_CriticalPathMode"

    shpKnob.Left = knobLeft
    shpKnob.Top = shpBg.Top + ((shpBg.Height - shpKnob.Height) / 2)

SafeExit:
End Sub


'------------------------------------------------------------------------------
' FR: Verifie ou cree Critical Path Mode Shapes si necessaire.
' EN: Ensures or creates Critical Path Mode Shapes when needed.
'------------------------------------------------------------------------------
Private Sub Ensure_CriticalPathMode_Shapes()

    Dim ws As Worksheet
    Dim labelCell As Range
    Dim stateCell As Range

    Dim shpLabel As Shape
    Dim shpBg As Shape
    Dim shpKnob As Shape

    Dim labelLeft As Double
    Dim labelTop As Double
    Dim labelWidth As Double
    Dim labelHeight As Double

    Dim trackWidth As Double
    Dim trackHeight As Double
    Dim knobSize As Double
    Dim trackLeft As Double
    Dim trackTop As Double
    Dim knobLeft As Double
    Dim knobTop As Double

    Dim isMulti As Boolean

    On Error GoTo SafeExit

    Set ws = ThisWorkbook.Worksheets(CP_MODE_WS)
    Set labelCell = ws.Range(CP_MODE_LABEL_CELL)
    Set stateCell = ws.Range(CP_MODE_STATE_CELL)

    'Nettoyage ancienne version + version avec textes Global/Multi.
    DeleteShapeIfExists_CriticalPathMode ws, CP_LABEL_SHAPE
    DeleteShapeIfExists_CriticalPathMode ws, CP_BG_SHAPE
    DeleteShapeIfExists_CriticalPathMode ws, CP_KNOB_SHAPE
    DeleteShapeIfExists_CriticalPathMode ws, "shp_WBS_CPMode_Left"
    DeleteShapeIfExists_CriticalPathMode ws, "shp_WBS_CPMode_Right"

    isMulti = IsCriticalPathMultiNetworkEnabled()

    labelLeft = labelCell.Left + 6
    labelTop = labelCell.Top + 3
    labelWidth = labelCell.Width + 14
    labelHeight = labelCell.Height - 6

    trackWidth = 42
    trackHeight = 20
    knobSize = 14

    trackLeft = stateCell.Left + ((stateCell.Width - trackWidth) / 2)
    trackTop = stateCell.Top + ((stateCell.Height - trackHeight) / 2)
    knobTop = trackTop + ((trackHeight - knobSize) / 2)

    If isMulti Then
        knobLeft = trackLeft + trackWidth - knobSize - 3
    Else
        knobLeft = trackLeft + 3
    End If

    Set shpLabel = ws.Shapes.AddTextbox( _
        msoTextOrientationHorizontal, _
        labelLeft, _
        labelTop, _
        labelWidth, _
        labelHeight)

    shpLabel.Name = CP_LABEL_SHAPE

    With shpLabel
        .Line.Visible = msoFalse
        .Fill.Visible = msoFalse
        .OnAction = ""
        .Placement = xlMoveAndSize
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .TextFrame2.MarginLeft = 0
        .TextFrame2.MarginRight = 0
        .TextFrame2.MarginTop = 0
        .TextFrame2.MarginBottom = 0
        .TextFrame2.TextRange.Text = "Multi Critical Path"
        .TextFrame2.TextRange.Font.Size = 9.5
        .TextFrame2.TextRange.Font.Bold = msoTrue
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignLeft
    End With

    Set shpBg = ws.Shapes.AddShape( _
        msoShapeRoundedRectangle, _
        trackLeft, _
        trackTop, _
        trackWidth, _
        trackHeight)

    shpBg.Name = CP_BG_SHAPE

    With shpBg
        .Adjustments.item(1) = 0.5
        .OnAction = "Toggle_CriticalPathMode"
        .Placement = xlMoveAndSize
        .Line.Weight = 1
        .Shadow.Visible = msoFalse
    End With

    Set shpKnob = ws.Shapes.AddShape( _
        msoShapeOval, _
        knobLeft, _
        knobTop, _
        knobSize, _
        knobSize)

    shpKnob.Name = CP_KNOB_SHAPE

    With shpKnob
        .OnAction = "Toggle_CriticalPathMode"
        .Placement = xlMoveAndSize
        .Fill.ForeColor.RGB = RGB(255, 255, 255)
        .Line.ForeColor.RGB = RGB(150, 150, 150)
        .Line.Weight = 0.75
        .Shadow.Visible = msoFalse
    End With

    Refresh_CriticalPathMode_Toggle_Visual

SafeExit:
End Sub

'------------------------------------------------------------------------------
' FR: Invalide l'etat incremental apres un changement Critical/Longest Path afin d'imposer un recalcul complet coherent.
' EN: Invalidates incremental state after a Critical/Longest Path change so the next calculation is coherently full.
'------------------------------------------------------------------------------

Private Sub ForceFullRecalcAfterCriticalPathModeChange()

    'On invalide le snapshot incrémental.
    'Get_Changed_TaskIds verra Run Status <> OK et forcera un full recalcul.

    SetCalcStateRunStatusDirty

End Sub

'------------------------------------------------------------------------------
' FR: Met a jour Calc State Run Status Dirty dans le contexte critical path mode.
' EN: Updates Calc State Run Status Dirty in the critical path mode context.
'------------------------------------------------------------------------------
Private Sub SetCalcStateRunStatusDirty()

    CalcState_MarkDirty

End Sub

'------------------------------------------------------------------------------
' FR: Formate Critical Path Mode Toggle Track pour l'affichage ou l'ecriture.
' EN: Formats Critical Path Mode Toggle Track for display or writing.
'------------------------------------------------------------------------------
Private Sub FormatCriticalPathModeToggleTrack( _
    ByVal shp As Shape, _
    ByVal leftPos As Double, _
    ByVal topPos As Double, _
    ByVal macroName As String, _
    ByVal isOn As Boolean)

    With shp
        .Left = leftPos
        .Top = topPos
        .Width = 27
        .Height = 12
        .Adjustments.item(1) = 0.5
        .OnAction = macroName
        .Placement = xlMoveAndSize
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
' FR: Formate Critical Path Mode Toggle Knob pour l'affichage ou l'ecriture.
' EN: Formats Critical Path Mode Toggle Knob for display or writing.
'------------------------------------------------------------------------------
Private Sub FormatCriticalPathModeToggleKnob( _
    ByVal shp As Shape, _
    ByVal leftPos As Double, _
    ByVal topPos As Double, _
    ByVal macroName As String)

    With shp
        .Left = leftPos
        .Top = topPos
        .Width = 9
        .Height = 9
        .OnAction = macroName
        .Placement = xlMoveAndSize
        .Fill.ForeColor.RGB = RGB(255, 255, 255)
        .Line.ForeColor.RGB = RGB(150, 150, 150)
        .Line.Weight = 0.75
        .Shadow.Visible = msoFalse
    End With

End Sub

'------------------------------------------------------------------------------
' FR: Formate Critical Path Mode Toggle Label pour l'affichage ou l'ecriture.
' EN: Formats Critical Path Mode Toggle Label for display or writing.
'------------------------------------------------------------------------------
Private Sub FormatCriticalPathModeToggleLabel( _
    ByVal shp As Shape, _
    ByVal leftPos As Double, _
    ByVal topPos As Double, _
    ByVal widthVal As Double, _
    ByVal heightVal As Double, _
    ByVal labelText As String, _
    ByVal macroName As String, _
    ByVal isOn As Boolean)

    With shp
        .Left = leftPos
        .Top = topPos
        .Width = widthVal
        .Height = heightVal
        .Line.Visible = msoFalse
        .Fill.Visible = msoFalse
        .OnAction = macroName
        .Placement = xlMoveAndSize
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .TextFrame2.MarginLeft = 0
        .TextFrame2.MarginRight = 0
        .TextFrame2.MarginTop = 0
        .TextFrame2.MarginBottom = 0
        .TextFrame2.TextRange.Text = labelText
        .TextFrame2.TextRange.Font.Size = 9.5
        .TextFrame2.TextRange.Font.Bold = msoTrue
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter

        If isOn Then
            .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(68, 114, 196)
        Else
            .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(90, 90, 90)
        End If
    End With

End Sub

'------------------------------------------------------------------------------
' FR: Supprime Shape If Exists Critical Path Mode du contexte critical path mode.
' EN: Removes Shape If Exists Critical Path Mode from the critical path mode context.
'------------------------------------------------------------------------------
Private Sub DeleteShapeIfExists_CriticalPathMode( _
    ByVal ws As Worksheet, _
    ByVal shapeName As String)

    Dim shp As Shape

    On Error Resume Next
    Set shp = ws.Shapes(shapeName)
    On Error GoTo 0

    If Not shp Is Nothing Then shp.Delete

End Sub

'------------------------------------------------------------------------------
' FR: Bascule l'etat Critical Path Mode Core et met a jour les sorties associees.
' EN: Toggles Critical Path Mode Core state and updates related outputs.
'------------------------------------------------------------------------------
Private Sub Toggle_CriticalPathMode_Core(ByVal refreshGanttAfterCalc As Boolean)

    Dim enabledNext As Boolean
    Dim oldScreenUpdating As Boolean
    Dim oldEvents As Boolean

    On Error GoTo ErrHandler

    oldScreenUpdating = Application.ScreenUpdating
    oldEvents = Application.EnableEvents

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    enabledNext = Not IsCriticalPathMultiNetworkEnabled()

    SetCriticalPathModeVisualState enabledNext
    Refresh_CriticalPathMode_Toggle_Visual
    GanttUiControls_RefreshCriticalPathMultiVisual ThisWorkbook.Worksheets("GANTT")

    ForceFullRecalcAfterCriticalPathModeChange

    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreenUpdating

    Run_Calc_Engine

    If refreshGanttAfterCalc Then
        If Not CommitGanttUpdate(GANTT_DATA_SOURCE_NORMAL, GANTT_UPDATE_SCOPE_INCREMENTAL, GANTT_RENDER_INTENT_SHOW, "Toggle_CriticalPathMode") Then
            Err.Raise 5, "Toggle_CriticalPathMode_Core", "Gantt render did not reach READY state after Critical Path mode change."
        End If
    End If

    Refresh_CriticalPathMode_Toggle_Visual
    GanttUiControls_RefreshCriticalPathMultiVisual ThisWorkbook.Worksheets("GANTT")

    Exit Sub

ErrHandler:
    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreenUpdating

    CalcBridge_ShowSingleConsoleMessage _
        "STOP", _
        "Erreur dans Toggle_CriticalPathMode : " & Err.Description, _
        "Error in Toggle_CriticalPathMode: " & Err.Description

End Sub

'------------------------------------------------------------------------------
' FR: Traite la reference Cleanup Critical Path Mode Gantt Legacy Shapes sans modifier les donnees d'entree.
' EN: Handles the Cleanup Critical Path Mode Gantt Legacy Shapes reference without mutating input data.
'------------------------------------------------------------------------------

Public Sub Cleanup_CriticalPathMode_Gantt_Legacy_Shapes()

    Dim ws As Worksheet

    On Error GoTo SafeExit

    Set ws = ThisWorkbook.Worksheets(CP_GANTT_WS)

    DeleteShapeIfExists_CriticalPathMode ws, "shp_GANTT_CPMode_Left"
    DeleteShapeIfExists_CriticalPathMode ws, "shp_GANTT_CPMode_Right"

    'Ne supprime pas les 3 shapes actuelles :
    'CP_GANTT_LABEL_SHAPE / CP_GANTT_BG_SHAPE / CP_GANTT_KNOB_SHAPE.
    'Elles sont maintenant repositionnées proprement par BuildFixedHeaderToggles.

SafeExit:
End Sub
