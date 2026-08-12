Attribute VB_Name = "mod_GanttGeometry"
Option Explicit

'===============================================================================
' MODULE : mod_GanttGeometry
' DOMAINE / DOMAIN : Gantt
'
' FR
' Fournit les calculs purs de geometrie et de positionnement du domaine.
' Ne cree aucune shape et ne decide aucun workflow.
'
' EN
' Provides pure domain geometry and positioning calculations.
' Creates no shape and decides no workflow.
'
' CONTRATS / CONTRACTS : GetGanttRowTop, GetGanttRowHeight, GetGanttRowMid, GetGanttBarTop, GetGanttBarHeight, GetGanttCompactTaskMarkerSize, GetGanttMilestoneSize, GetGanttSummaryTop
' CALLBACKS EXTERNES / EXTERNAL CALLBACKS : Aucun / None
'===============================================================================


' Pure vertical geometry and render placement helpers extracted from mod_Gantt.
' Keep offsets and placement semantics identical to the legacy renderer.
'=====================================================

Private Const FIRST_TIMELINE_COL As Long = 11

'------------------------------------------------------------------------------
' FR: Retourne une valeur d'etat ou une reference utilisee par le rendu GANTT.
' EN: Returns a state value or reference used by GANTT rendering.
'------------------------------------------------------------------------------
Public Function GetGanttRowTop(ByVal ws As Worksheet, ByVal ganttRow As Long) As Double

    GetGanttRowTop = ws.cells(ganttRow, FIRST_TIMELINE_COL).Top

End Function

'------------------------------------------------------------------------------
' FR: Retourne une valeur d'etat ou une reference utilisee par le rendu GANTT.
' EN: Returns a state value or reference used by GANTT rendering.
'------------------------------------------------------------------------------
Public Function GetGanttRowHeight(ByVal ws As Worksheet, ByVal ganttRow As Long) As Double

    GetGanttRowHeight = ws.rows(ganttRow).Height

End Function

'------------------------------------------------------------------------------
' FR: Retourne une valeur d'etat ou une reference utilisee par le rendu GANTT.
' EN: Returns a state value or reference used by GANTT rendering.
'------------------------------------------------------------------------------
Public Function GetGanttRowMid(ByVal ws As Worksheet, ByVal ganttRow As Long) As Double

    GetGanttRowMid = GetGanttRowTop(ws, ganttRow) + (GetGanttRowHeight(ws, ganttRow) / 2)

End Function

'------------------------------------------------------------------------------
' FR: Retourne une valeur d'etat ou une reference utilisee par le rendu GANTT.
' EN: Returns a state value or reference used by GANTT rendering.
'------------------------------------------------------------------------------
Public Function GetGanttBarTop(ByVal ws As Worksheet, ByVal ganttRow As Long) As Double

    GetGanttBarTop = GetGanttRowTop(ws, ganttRow) + 4

End Function

'------------------------------------------------------------------------------
' FR: Retourne une valeur d'etat ou une reference utilisee par le rendu GANTT.
' EN: Returns a state value or reference used by GANTT rendering.
'------------------------------------------------------------------------------
Public Function GetGanttBarHeight(ByVal ws As Worksheet, ByVal ganttRow As Long) As Double

    GetGanttBarHeight = GetGanttRowHeight(ws, ganttRow) - 8
    If GetGanttBarHeight < 2 Then GetGanttBarHeight = 2

End Function

'------------------------------------------------------------------------------
' FR: Retourne une valeur d'etat ou une reference utilisee par le rendu GANTT.
' EN: Returns a state value or reference used by GANTT rendering.
'------------------------------------------------------------------------------
Public Function GetGanttCompactTaskMarkerSize( _
    ByVal ws As Worksheet, _
    ByVal ganttRow As Long, _
    ByVal cellWidth As Double) As Double

    GetGanttCompactTaskMarkerSize = WorksheetFunction.Min(GetGanttBarHeight(ws, ganttRow), cellWidth - 6)
    If GetGanttCompactTaskMarkerSize < 2 Then GetGanttCompactTaskMarkerSize = 2

End Function
'------------------------------------------------------------------------------
' FR: Retourne une valeur d'etat ou une reference utilisee par le rendu GANTT.
' EN: Returns a state value or reference used by GANTT rendering.
'------------------------------------------------------------------------------
Public Function GetGanttMilestoneSize(ByVal ws As Worksheet, ByVal ganttRow As Long) As Double

    GetGanttMilestoneSize = GetGanttRowHeight(ws, ganttRow) - 6
    If GetGanttMilestoneSize < 2 Then GetGanttMilestoneSize = 2

End Function

'------------------------------------------------------------------------------
' FR: Retourne une valeur d'etat ou une reference utilisee par le rendu GANTT.
' EN: Returns a state value or reference used by GANTT rendering.
'------------------------------------------------------------------------------
Public Function GetGanttSummaryTop(ByVal ws As Worksheet, ByVal ganttRow As Long) As Double

    GetGanttSummaryTop = GetGanttRowTop(ws, ganttRow) + 4

End Function

'------------------------------------------------------------------------------
' FR: Retourne une valeur d'etat ou une reference utilisee par le rendu GANTT.
' EN: Returns a state value or reference used by GANTT rendering.
'------------------------------------------------------------------------------
Public Function GetGanttSummaryBottom(ByVal ws As Worksheet, ByVal ganttRow As Long) As Double

    GetGanttSummaryBottom = GetGanttRowTop(ws, ganttRow) + GetGanttRowHeight(ws, ganttRow) - 4

End Function

'------------------------------------------------------------------------------
' FR: Actualise Apply Gantt Render Shape Placement sans modifier les regles metier qui produisent les donnees.
' EN: Refreshes Apply Gantt Render Shape Placement without changing the business rules that produce the data.
'------------------------------------------------------------------------------

Public Sub ApplyGanttRenderShapePlacement(ByVal shp As Shape)

    On Error Resume Next

    'Important:
    'The Gantt renderer deletes/recreates shapes on each refresh.
    'Using xlMoveAndSize lets Excel re-anchor and micro-resize shapes when
    'columns/rows/merged headers are rebuilt, which can create cumulative drift.
    'Keep rendered shapes free-floating; their exact Left/Top/Width/Height are
    'fully controlled by the renderer.
    shp.Placement = xlMoveAndSize

    On Error GoTo 0

End Sub

'------------------------------------------------------------------------------
' FR: Actualise Apply Gantt Render Line Placement sans modifier les regles metier qui produisent les donnees.
' EN: Refreshes Apply Gantt Render Line Placement without changing the business rules that produce the data.
'------------------------------------------------------------------------------

Public Sub ApplyGanttRenderLinePlacement(ByVal shp As Shape)

    On Error Resume Next

    'Same rule as bars/milestones:
    'dependency lines and Today line are redrawn from scratch, so they must not
    'be cell-resized by Excel between layout rebuild and final render.
    shp.Placement = xlMoveAndSize

    On Error GoTo 0

End Sub
