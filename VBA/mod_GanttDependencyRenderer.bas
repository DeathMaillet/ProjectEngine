Attribute VB_Name = "mod_GanttDependencyRenderer"
Option Explicit

'===============================================================================
' MODULE : mod_GanttDependencyRenderer
' DOMAINE / DOMAIN : Gantt
'
' FR
' Construit le rendu Excel du composant a partir de donnees deja preparees.
' Ne decide pas les donnees metier a calculer.
'
' EN
' Builds the component's Excel rendering from prepared data.
' Does not decide business data to calculate.
'
' CONTRATS / CONTRACTS : GanttDependency_ClearExpandedLinksCache, DrawDependencyLinks, GetTaskMidX
' CALLBACKS EXTERNES / EXTERNAL CALLBACKS : Aucun / None
'===============================================================================

Private Const GANTT_SHEET As String = "GANTT"
Private Const WBS_SHEET As String = "WBS"
Private Const WBS_TABLE As String = "tbl_WBS"
Private Const CALC_SHEET As String = "CALC"
Private Const CALC_TABLE As String = "tbl_CALC"

Private Const FIRST_TIMELINE_COL As Long = 11
Private Const HEADER_ROW_1 As Long = 3
Private Const HEADER_ROW_2 As Long = 4
Private Const FIRST_TASK_ROW As Long = 5
Private Const COL_WBS As Long = 1
Private Const COL_TASK As Long = 2

Private Const COLOR_TASK_BLUE As Long = 12874308
Private Const COLOR_TASK_CRITICAL As Long = 192
Private Const COLOR_PROGRESS_GREEN As Long = 4699504
Private Const COLOR_PROGRESS_ORANGE As Long = 3248093

Private Const GANTT_VIEW_DETAIL As String = "DETAIL"
Private Const GANTT_VIEW_SUMMARY As String = "SUMMARY"
Private Const GANTT_ANALYTICS_PATH_NONE As String = "NONE"
Private Const GANTT_ANALYTICS_PATH_CP As String = "CP"
Private Const GANTT_ANALYTICS_PATH_LP As String = "LP"

Private Const LINK_STUB As Double = 6
Private Const LINK_MIN_CHANNEL_GAP As Double = 4
Private Const LINK_EDGE_PADDING As Double = 8
Private Const LINK_ANCHOR_START As String = "START"
Private Const LINK_ANCHOR_FINISH As String = "FINISH"

Private gExpandedLinks As Object

'------------------------------------------------------------------------------
' FR: Reinitialise Gantt Dependency Clear Expanded Links Cache dans le perimetre possede par le composant.
' EN: Resets Gantt Dependency Clear Expanded Links Cache within the state owned by the component.
'------------------------------------------------------------------------------

Public Sub GanttDependency_ClearExpandedLinksCache()
    Set gExpandedLinks = Nothing
End Sub



'------------------------------------------------------------------------------
' FR:
' Dessine les liens de dependance depuis le cache tbl_LOGIC_LINKS expanse,
' en sautant les modes Week/Month.
'
' EN:
' Draws dependency links from the expanded tbl_LOGIC_LINKS cache,
' skipping Week/Month aggregated modes.
'
' Entrees / Inputs:
' - data WBS, rowById, hasChildren, maps base/test, mode TEST et timeline Day.
'
' Sorties / Outputs:
' - Shapes DEP_* composees de segments avec fleches.
'
' Appele par / Called by:
' - RunGanttRefreshCore.
'
' Notes:
' - Fortement couple a CALC/tbl_LOGIC_LINKS et a la geometrie des task bars.
'------------------------------------------------------------------------------
Public Sub DrawDependencyLinks( _
    ByVal wsGantt As Worksheet, _
    ByVal mapWBS As Object, _
    ByRef dataArr As Variant, _
    ByVal hasChildren As Object, _
    ByVal rowById As Object, _
    ByVal projectStart As Variant, _
    ByVal totalDays As Long, _
    ByVal baseById As Object, _
    ByVal testById As Object, _
    ByVal isTestMode As Boolean)

    Dim perfScope As clsPerfScope

    Dim succId As Variant
    Dim linkItem As Variant
    Dim predId As String
    Dim shapePrefix As String
    Dim linkIndex As Long
    Dim anchorCache As Object

    Set perfScope = Profiler_BeginScope("DrawDependencyLinks", "Dependency Render")

    If IsAggregatedScaleMode() Then Exit Sub

    On Error GoTo SafeExit

    EnsureExpandedLinksCacheFromCalc
    If Not HasExpandedLinksAvailable() Then Exit Sub

    Set anchorCache = CreateObject("Scripting.Dictionary")

    For Each succId In gExpandedLinks.Keys

        linkIndex = 0

        For Each linkItem In gExpandedLinks(CStr(succId))

            predId = Trim$(CStr(linkItem("PredID")))

            If predId <> "" Then
                If rowById.Exists(predId) And rowById.Exists(CStr(succId)) Then

                    '==================================================
                    ' NOUVELLE RÈGLE :
                    ' on accepte aussi l'affichage des liens parent -> parent
                    ' si le lien existe dans CALC, on le dessine
                    '==================================================
                    linkIndex = linkIndex + 1
                    shapePrefix = "DEP_" & predId & "_" & CStr(succId) & "_" & CStr(linkIndex)

                    DrawSingleDependencyLink _
                        wsGantt, mapWBS, dataArr, hasChildren, rowById, _
                        projectStart, totalDays, _
                        predId, CStr(succId), _
                        baseById, testById, isTestMode, _
                        anchorCache, shapePrefix, _
                        GetLinkTypeFromItem(linkItem), _
                        GetLinkLagFromItem(linkItem)

                End If
            End If

NextLink:
        Next linkItem

    Next succId

SafeExit:
End Sub
'------------------------------------------------------------------------------
' FR: Calcule ou dessine une partie des liens de dependance visibles dans le GANTT.
' EN: Computes or draws part of the dependency links visible in GANTT.
'------------------------------------------------------------------------------
Private Function GetLinkTypeFromItem(ByVal linkItem As Variant) As String

    On Error GoTo SafeExit

    If IsObject(linkItem) Then
        If linkItem.Exists("LinkType") Then
            GetLinkTypeFromItem = UCase$(Trim$(CStr(linkItem("LinkType"))))
            If GetLinkTypeFromItem = "" Then GetLinkTypeFromItem = "FS"
            Exit Function
        End If
    End If

SafeExit:
    GetLinkTypeFromItem = "FS"

End Function
'------------------------------------------------------------------------------
' FR: Calcule ou dessine une partie des liens de dependance visibles dans le GANTT.
' EN: Computes or draws part of the dependency links visible in GANTT.
'------------------------------------------------------------------------------
Private Function GetLinkLagFromItem(ByVal linkItem As Variant) As Double

    On Error GoTo SafeExit

    If IsObject(linkItem) Then
        If linkItem.Exists("Lag") Then
            If IsNumeric(linkItem("Lag")) Then
                GetLinkLagFromItem = CDbl(linkItem("Lag"))
                Exit Function
            End If
        End If
    End If

SafeExit:
    GetLinkLagFromItem = 0#

End Function
'------------------------------------------------------------------------------
' FR:
' Route un lien FS/SS/FF entre deux taches en choisissant les ancres,
' les points d'entree et le chemin visuel adapte au lag.
'
' EN:
' Routes one FS/SS/FF link between two tasks by choosing anchors,
' entry points, and a visual path suited to lag.
'
' Entrees / Inputs:
' - IDs pred/succ, type de lien, lag, cache d'ancres, maps base/test.
'
' Sorties / Outputs:
' - Segments DEP_* horizontaux/verticaux formant le connecteur.
'
' Appele par / Called by:
' - DrawDependencyLinks.
'
' Notes:
' - Zone a ne pas refactoriser brutalement: nombreuses regles visuelles FS same-day/negative.
'------------------------------------------------------------------------------
Private Sub DrawSingleDependencyLink( _
    ByVal wsGantt As Worksheet, _
    ByVal mapWBS As Object, _
    ByRef dataArr As Variant, _
    ByVal hasChildren As Object, _
    ByVal rowById As Object, _
    ByVal projectStart As Variant, _
    ByVal totalDays As Long, _
    ByVal predId As String, _
    ByVal succId As String, _
    ByVal baseById As Object, _
    ByVal testById As Object, _
    ByVal isTestMode As Boolean, _
    ByVal anchorCache As Object, _
    ByVal shapePrefix As String, _
    ByVal linkType As String, _
    ByVal linkLag As Double)

    Dim perfScope As clsPerfScope

    Dim predRow As Long
    Dim succRow As Long
    Dim predDataRow As Long
    Dim succDataRow As Long

    Dim predX As Double
    Dim predY As Double
    Dim succX As Double
    Dim succY As Double

    Dim succTopX As Double
    Dim succTopY As Double
    Dim succMidLeftX As Double
    Dim succMidLeftY As Double

    Dim predAnchorType As String
    Dim succAnchorType As String
    Dim predDate As Variant
    Dim succDate As Variant
    Dim gapDays As Long
    Dim useMidLeftEntry As Boolean

    Set perfScope = Profiler_BeginScope("DrawSingleDependencyLink", "Dependency Render")

    If Not rowById.Exists(predId) Then Exit Sub
    If Not rowById.Exists(succId) Then Exit Sub

    predRow = CLng(rowById(predId))
    succRow = CLng(rowById(succId))

    predDataRow = predRow - FIRST_TASK_ROW + 1
    succDataRow = succRow - FIRST_TASK_ROW + 1

    If predDataRow < 1 Or predDataRow > UBound(dataArr, 1) Then Exit Sub
    If succDataRow < 1 Or succDataRow > UBound(dataArr, 1) Then Exit Sub

    linkType = UCase$(Trim$(linkType))
    If linkType = "" Then linkType = "FS"

    GetLinkAnchorTypes linkType, predAnchorType, succAnchorType

    predDate = GetLinkReferenceDate(predId, predAnchorType, baseById, testById, isTestMode)
    succDate = GetLinkReferenceDate(succId, succAnchorType, baseById, testById, isTestMode)

    If Not HasValue(predDate) Then Exit Sub
    If Not HasValue(succDate) Then Exit Sub

    GetCachedTaskAnchorPointByType anchorCache, wsGantt, mapWBS, dataArr, hasChildren, projectStart, totalDays, predDataRow, _
        predAnchorType, predX, predY, baseById, testById, isTestMode

    Select Case linkType

        Case "SS"
            GetCachedTaskAnchorPointByType anchorCache, wsGantt, mapWBS, dataArr, hasChildren, projectStart, totalDays, succDataRow, _
                succAnchorType, succX, succY, baseById, testById, isTestMode

            gapDays = CLng(CDbl(succDate) - CDbl(predDate) - linkLag)
            RouteDependencyLink_SS wsGantt, shapePrefix, predX, predY, succX, succY, gapDays

        Case "FF"
            GetCachedTaskAnchorPointByType anchorCache, wsGantt, mapWBS, dataArr, hasChildren, projectStart, totalDays, succDataRow, _
                succAnchorType, succX, succY, baseById, testById, isTestMode

            gapDays = CLng(CDbl(succDate) - CDbl(predDate) - linkLag)
            RouteDependencyLink_FF wsGantt, shapePrefix, predX, predY, succX, succY, gapDays

        Case Else   ' FS
            gapDays = CLng(CDbl(succDate) - CDbl(predDate) - 1 - linkLag)

            ' On calcule les 2 points candidats côté successeur
            GetCachedTaskTopEntryPoint anchorCache, wsGantt, mapWBS, dataArr, projectStart, totalDays, succDataRow, _
                succTopX, succTopY, baseById, testById, isTestMode

            GetCachedTaskStartMidEntryPoint anchorCache, wsGantt, mapWBS, dataArr, hasChildren, projectStart, totalDays, succDataRow, _
                succMidLeftX, succMidLeftY, baseById, testById, isTestMode

            ' Règle corrigée :
            ' on ne décide PAS avec gapDays=0/1
            ' on décide avec la place horizontale réelle entre pred et l'entrée gauche du successeur
            useMidLeftEntry = HasRoomForFsMidLeftEntry(predX, succMidLeftX, wsGantt.cells(HEADER_ROW_2, FIRST_TIMELINE_COL).Width)

            If gapDays < 0 Then
                succX = succMidLeftX
                succY = succMidLeftY
                RouteDependencyLink_FS_Negative wsGantt, shapePrefix, predX, predY, succX, succY, gapDays

            ElseIf useMidLeftEntry Then
                succX = succMidLeftX
                succY = succMidLeftY
                RouteDependencyLink_FS_Normal wsGantt, shapePrefix, predX, predY, succX, succY, gapDays

            Else
                succX = succTopX
                succY = succTopY
                RouteDependencyLink_FS_SameDay wsGantt, shapePrefix, predX, predY, succX, succY
            End If

    End Select

End Sub
'------------------------------------------------------------------------------
' FR: Execute le helper Get Cached Task Anchor Point By Type dans le workflow de rendu GANTT.
' EN: Runs the Get Cached Task Anchor Point By Type helper in the GANTT rendering workflow.
'------------------------------------------------------------------------------
Private Sub GetCachedTaskAnchorPointByType( _
    ByVal anchorCache As Object, _
    ByVal ws As Worksheet, _
    ByVal mapWBS As Object, _
    ByRef dataArr As Variant, _
    ByVal hasChildren As Object, _
    ByVal projectStart As Variant, _
    ByVal totalDays As Long, _
    ByVal dataRow As Long, _
    ByVal anchorType As String, _
    ByRef xOut As Double, _
    ByRef yOut As Double, _
    ByVal baseById As Object, _
    ByVal testById As Object, _
    ByVal isTestMode As Boolean)

    Dim cacheKey As String
    Dim cachedValue As Variant

    cacheKey = "TYPE|" & CStr(dataRow) & "|" & UCase$(Trim$(anchorType))

    If Not anchorCache Is Nothing Then
        If anchorCache.Exists(cacheKey) Then
            cachedValue = anchorCache(cacheKey)
            xOut = CDbl(cachedValue(0))
            yOut = CDbl(cachedValue(1))
            Exit Sub
        End If
    End If

    GetTaskAnchorPointByType ws, mapWBS, dataArr, hasChildren, projectStart, totalDays, dataRow, _
        anchorType, xOut, yOut, baseById, testById, isTestMode

    If Not anchorCache Is Nothing Then anchorCache(cacheKey) = Array(xOut, yOut)

End Sub
'------------------------------------------------------------------------------
' FR: Execute le helper Get Cached Task Top Entry Point dans le workflow de rendu GANTT.
' EN: Runs the Get Cached Task Top Entry Point helper in the GANTT rendering workflow.
'------------------------------------------------------------------------------
Private Sub GetCachedTaskTopEntryPoint( _
    ByVal anchorCache As Object, _
    ByVal ws As Worksheet, _
    ByVal mapWBS As Object, _
    ByRef dataArr As Variant, _
    ByVal projectStart As Variant, _
    ByVal totalDays As Long, _
    ByVal dataRow As Long, _
    ByRef xOut As Double, _
    ByRef yOut As Double, _
    ByVal baseById As Object, _
    ByVal testById As Object, _
    ByVal isTestMode As Boolean)

    Dim cacheKey As String
    Dim cachedValue As Variant

    cacheKey = "TOP|" & CStr(dataRow)

    If Not anchorCache Is Nothing Then
        If anchorCache.Exists(cacheKey) Then
            cachedValue = anchorCache(cacheKey)
            xOut = CDbl(cachedValue(0))
            yOut = CDbl(cachedValue(1))
            Exit Sub
        End If
    End If

    GetTaskTopEntryPoint ws, mapWBS, dataArr, projectStart, totalDays, dataRow, _
        xOut, yOut, baseById, testById, isTestMode

    If Not anchorCache Is Nothing Then anchorCache(cacheKey) = Array(xOut, yOut)

End Sub
'------------------------------------------------------------------------------
' FR: Execute le helper Get Cached Task Start Mid Entry Point dans le workflow de rendu GANTT.
' EN: Runs the Get Cached Task Start Mid Entry Point helper in the GANTT rendering workflow.
'------------------------------------------------------------------------------
Private Sub GetCachedTaskStartMidEntryPoint( _
    ByVal anchorCache As Object, _
    ByVal ws As Worksheet, _
    ByVal mapWBS As Object, _
    ByRef dataArr As Variant, _
    ByVal hasChildren As Object, _
    ByVal projectStart As Variant, _
    ByVal totalDays As Long, _
    ByVal dataRow As Long, _
    ByRef xOut As Double, _
    ByRef yOut As Double, _
    ByVal baseById As Object, _
    ByVal testById As Object, _
    ByVal isTestMode As Boolean)

    Dim cacheKey As String
    Dim cachedValue As Variant

    cacheKey = "MIDLEFT|" & CStr(dataRow)

    If Not anchorCache Is Nothing Then
        If anchorCache.Exists(cacheKey) Then
            cachedValue = anchorCache(cacheKey)
            xOut = CDbl(cachedValue(0))
            yOut = CDbl(cachedValue(1))
            Exit Sub
        End If
    End If

    GetTaskStartMidEntryPoint ws, mapWBS, dataArr, hasChildren, projectStart, totalDays, dataRow, _
        xOut, yOut, baseById, testById, isTestMode

    If Not anchorCache Is Nothing Then anchorCache(cacheKey) = Array(xOut, yOut)

End Sub
'------------------------------------------------------------------------------
' FR: Retourne une decision de rendu ou d'etat utilisee par le workflow GANTT.
' EN: Returns a rendering or state decision used by the GANTT workflow.
'------------------------------------------------------------------------------
Private Function HasRoomForFsMidLeftEntry( _
    ByVal predX As Double, _
    ByVal succMidLeftX As Double, _
    ByVal cellWidth As Double) As Boolean

    Dim minNeeded As Double

    ' Il faut une vraie place visuelle pour arriver par la gauche.
    ' gapDays = 0 peut quand même avoir assez de place à l’écran.
    minNeeded = WorksheetFunction.Max(10, cellWidth * 0.55)

    HasRoomForFsMidLeftEntry = ((succMidLeftX - predX) >= minNeeded)

End Function
'------------------------------------------------------------------------------
' FR: Calcule ou dessine une partie des liens de dependance visibles dans le GANTT.
' EN: Computes or draws part of the dependency links visible in GANTT.
'------------------------------------------------------------------------------
Private Sub RouteDependencyLink_FS_SameDay( _
    ByVal wsGantt As Worksheet, _
    ByVal shapePrefix As String, _
    ByVal predX As Double, _
    ByVal predY As Double, _
    ByVal succX As Double, _
    ByVal succY As Double)

    DrawLinkSegment wsGantt, shapePrefix & "_1", predX, predY, succX, predY, False
    DrawLinkSegment wsGantt, shapePrefix & "_2", succX, predY, succX, succY, True

End Sub
'------------------------------------------------------------------------------
' FR: Execute le helper Get Task Start Mid Entry Point dans le workflow de rendu GANTT.
' EN: Runs the Get Task Start Mid Entry Point helper in the GANTT rendering workflow.
'------------------------------------------------------------------------------
Private Sub GetTaskStartMidEntryPoint( _
    ByVal ws As Worksheet, _
    ByVal mapWBS As Object, _
    ByRef dataArr As Variant, _
    ByVal hasChildren As Object, _
    ByVal projectStart As Variant, _
    ByVal totalDays As Long, _
    ByVal dataRow As Long, _
    ByRef xOut As Double, _
    ByRef yOut As Double, _
    ByVal baseById As Object, _
    ByVal testById As Object, _
    ByVal isTestMode As Boolean)

    ' Entrée milieu gauche = vrai point milieu côté gauche
    GetTaskAnchorPointBySide ws, mapWBS, dataArr, hasChildren, projectStart, totalDays, dataRow, _
        "LEFT", xOut, yOut, baseById, testById, isTestMode

End Sub
'------------------------------------------------------------------------------
' FR: Calcule ou dessine une partie des liens de dependance visibles dans le GANTT.
' EN: Computes or draws part of the dependency links visible in GANTT.
'------------------------------------------------------------------------------
Private Sub RouteDependencyLink_FS_Normal( _
    ByVal wsGantt As Worksheet, _
    ByVal shapePrefix As String, _
    ByVal predX As Double, _
    ByVal predY As Double, _
    ByVal succX As Double, _
    ByVal succY As Double, _
    ByVal gapDays As Long)

    Dim endX As Double
    Dim directEnough As Boolean
    Dim routeAbove As Boolean
    Dim laneY As Double
    Dim bendX As Double
    Dim entryGap As Double
    Dim finalX As Double
    Dim cellWidth As Double

    cellWidth = wsGantt.cells(HEADER_ROW_2, FIRST_TIMELINE_COL).Width

    ' IMPORTANT : plus aucun cas spécial ici basé sur gapDays = 0.
    ' Le choix top / milieu-gauche a déjà été fait en amont.

    If gapDays <= 1 Then
        entryGap = 8
    Else
        entryGap = 8 + (cellWidth / 2)
    End If

    endX = succX - entryGap
    If endX <= predX + 4 Then endX = succX - 4
    If endX <= predX + 2 Then endX = succX

    directEnough = (endX - predX >= LINK_STUB * 2)

    If directEnough Then
        If gapDays > 1 Then
            bendX = predX + LINK_STUB + (cellWidth / 3)
        Else
            bendX = predX + LINK_STUB
        End If

        If bendX > endX - LINK_STUB Then
            bendX = predX + ((endX - predX) / 2)
        End If

        finalX = succX - 2

        If finalX <= endX + 2 Then
            DrawLinkSegment wsGantt, shapePrefix & "_1", predX, predY, bendX, predY, False
            DrawLinkSegment wsGantt, shapePrefix & "_2", bendX, predY, bendX, succY, False
            DrawLinkSegment wsGantt, shapePrefix & "_3", bendX, succY, succX, succY, True
        Else
            DrawLinkSegment wsGantt, shapePrefix & "_1", predX, predY, bendX, predY, False
            DrawLinkSegment wsGantt, shapePrefix & "_2", bendX, predY, bendX, succY, False
            DrawLinkSegment wsGantt, shapePrefix & "_3", bendX, succY, endX, succY, False
            DrawLinkSegment wsGantt, shapePrefix & "_4", endX, succY, succX, succY, True
        End If

        Exit Sub
    End If

    routeAbove = (succY <= predY)

    If routeAbove Then
        laneY = WorksheetFunction.Min(predY, succY) - LINK_MIN_CHANNEL_GAP
    Else
        laneY = WorksheetFunction.Max(predY, succY) + LINK_MIN_CHANNEL_GAP
    End If

    If gapDays > 1 Then
        bendX = predX + LINK_STUB + (cellWidth / 3)
    Else
        bendX = predX + LINK_STUB
    End If

    If bendX >= succX - 6 Then
        bendX = predX + ((succX - predX) / 2)
    End If

    finalX = succX - 2

    If finalX <= endX + 2 Then
        DrawLinkSegment wsGantt, shapePrefix & "_1", predX, predY, bendX, predY, False
        DrawLinkSegment wsGantt, shapePrefix & "_2", bendX, predY, bendX, laneY, False
        DrawLinkSegment wsGantt, shapePrefix & "_3", bendX, laneY, succX - 2, laneY, False
        DrawLinkSegment wsGantt, shapePrefix & "_4", succX - 2, laneY, succX - 2, succY, False
        DrawLinkSegment wsGantt, shapePrefix & "_5", succX - 2, succY, succX, succY, True
    Else
        DrawLinkSegment wsGantt, shapePrefix & "_1", predX, predY, bendX, predY, False
        DrawLinkSegment wsGantt, shapePrefix & "_2", bendX, predY, bendX, laneY, False
        DrawLinkSegment wsGantt, shapePrefix & "_3", bendX, laneY, endX, laneY, False
        DrawLinkSegment wsGantt, shapePrefix & "_4", endX, laneY, endX, succY, False
        DrawLinkSegment wsGantt, shapePrefix & "_5", endX, succY, succX, succY, True
    End If

End Sub
'------------------------------------------------------------------------------
' FR: Calcule ou dessine une partie des liens de dependance visibles dans le GANTT.
' EN: Computes or draws part of the dependency links visible in GANTT.
'------------------------------------------------------------------------------
Private Sub RouteDependencyLink_FS_Negative( _
    ByVal wsGantt As Worksheet, _
    ByVal shapePrefix As String, _
    ByVal predX As Double, _
    ByVal predY As Double, _
    ByVal succX As Double, _
    ByVal succY As Double, _
    ByVal gapDays As Long)

    Dim cellWidth As Double
    Dim laneY As Double
    Dim leftX As Double

    cellWidth = wsGantt.cells(HEADER_ROW_2, FIRST_TIMELINE_COL).Width

    If succY <= predY Then
        laneY = WorksheetFunction.Min(predY, succY) - LINK_MIN_CHANNEL_GAP
    Else
        laneY = WorksheetFunction.Max(predY, succY) + LINK_MIN_CHANNEL_GAP
    End If

    leftX = WorksheetFunction.Min(predX, succX) - WorksheetFunction.Max(8, cellWidth / 2)
    leftX = leftX - WorksheetFunction.Max(6, Abs(gapDays) * (cellWidth / 2))

    DrawLinkSegment wsGantt, shapePrefix & "_1", predX, predY, predX, laneY, False
    DrawLinkSegment wsGantt, shapePrefix & "_2", predX, laneY, leftX, laneY, False
    DrawLinkSegment wsGantt, shapePrefix & "_3", leftX, laneY, leftX, succY, False
    DrawLinkSegment wsGantt, shapePrefix & "_4", leftX, succY, succX, succY, True

End Sub
'------------------------------------------------------------------------------
' FR: Calcule ou dessine une partie des liens de dependance visibles dans le GANTT.
' EN: Computes or draws part of the dependency links visible in GANTT.
'------------------------------------------------------------------------------
Private Sub RouteDependencyLink_SS( _
    ByVal wsGantt As Worksheet, _
    ByVal shapePrefix As String, _
    ByVal predX As Double, _
    ByVal predY As Double, _
    ByVal succX As Double, _
    ByVal succY As Double, _
    ByVal gapDays As Long)

    Dim cellWidth As Double
    Dim busX As Double

    cellWidth = wsGantt.cells(HEADER_ROW_2, FIRST_TIMELINE_COL).Width

    busX = WorksheetFunction.Min(predX, succX) - WorksheetFunction.Max(8, cellWidth / 2)

    If gapDays < 0 Then
        busX = busX - WorksheetFunction.Max(6, Abs(gapDays) * (cellWidth / 2))
    End If

    DrawLinkSegment wsGantt, shapePrefix & "_1", predX, predY, busX, predY, False
    DrawLinkSegment wsGantt, shapePrefix & "_2", busX, predY, busX, succY, False
    DrawLinkSegment wsGantt, shapePrefix & "_3", busX, succY, succX, succY, True

End Sub
'------------------------------------------------------------------------------
' FR: Calcule ou dessine une partie des liens de dependance visibles dans le GANTT.
' EN: Computes or draws part of the dependency links visible in GANTT.
'------------------------------------------------------------------------------
Private Sub RouteDependencyLink_FF( _
    ByVal wsGantt As Worksheet, _
    ByVal shapePrefix As String, _
    ByVal predX As Double, _
    ByVal predY As Double, _
    ByVal succX As Double, _
    ByVal succY As Double, _
    ByVal gapDays As Long)

    Dim cellWidth As Double
    Dim busX As Double

    cellWidth = wsGantt.cells(HEADER_ROW_2, FIRST_TIMELINE_COL).Width

    busX = WorksheetFunction.Min(predX, succX) - WorksheetFunction.Max(8, cellWidth / 2)

    If gapDays < 0 Then
        busX = busX - WorksheetFunction.Max(6, Abs(gapDays) * (cellWidth / 2))
    End If

    DrawLinkSegment wsGantt, shapePrefix & "_1", predX, predY, busX, predY, False
    DrawLinkSegment wsGantt, shapePrefix & "_2", busX, predY, busX, succY, False
    DrawLinkSegment wsGantt, shapePrefix & "_3", busX, succY, succX, succY, True

End Sub
'------------------------------------------------------------------------------
' FR: Execute le helper Get Task Top Entry Point dans le workflow de rendu GANTT.
' EN: Runs the Get Task Top Entry Point helper in the GANTT rendering workflow.
'------------------------------------------------------------------------------
Private Sub GetTaskTopEntryPoint( _
    ByVal ws As Worksheet, _
    ByVal mapWBS As Object, _
    ByRef dataArr As Variant, _
    ByVal projectStart As Variant, _
    ByVal totalDays As Long, _
    ByVal dataRow As Long, _
    ByRef xOut As Double, _
    ByRef yOut As Double, _
    ByVal baseById As Object, _
    ByVal testById As Object, _
    ByVal isTestMode As Boolean)

    Dim ganttRow As Long
    Dim idVal As String
    Dim startVal As Variant
    Dim finishVal As Variant
    Dim durationVal As Double
    Dim timelineLeftBound As Double
    Dim timelineRightBound As Double
    Dim topEntryOffset As Double

    ganttRow = FIRST_TASK_ROW + dataRow - 1
    idVal = Trim$(CStr(dataArr(dataRow, mapWBS("ID"))))

    startVal = GetRenderStartForCurrentScale(GanttLive_GetDisplayStart(idVal, baseById, testById, isTestMode))
    finishVal = GetRenderFinishForCurrentScale(GanttLive_GetDisplayFinish(idVal, baseById, testById, isTestMode))

    If Not HasValue(startVal) Or Not HasValue(finishVal) Then Exit Sub

    durationVal = CDbl(finishVal) - CDbl(startVal) + 1

    timelineLeftBound = ws.cells(HEADER_ROW_2, FIRST_TIMELINE_COL).Left + LINK_EDGE_PADDING
    timelineRightBound = ws.cells(HEADER_ROW_2, FIRST_TIMELINE_COL + totalDays - 1).Left + _
                         ws.cells(HEADER_ROW_2, FIRST_TIMELINE_COL + totalDays - 1).Width - LINK_EDGE_PADDING

    If durationVal <= 1 Then
        xOut = GetTaskMidX(ws, projectStart, startVal)
        yOut = GetGanttRowTop(ws, ganttRow) + 3
    Else
        xOut = TimelineLeft(ws, projectStart, startVal) + 4
        yOut = GetGanttBarTop(ws, ganttRow)
    End If

    topEntryOffset = ws.cells(HEADER_ROW_2, FIRST_TIMELINE_COL).Width * 0.15
    xOut = xOut + topEntryOffset

    If xOut < timelineLeftBound Then xOut = timelineLeftBound
    If xOut > timelineRightBound Then xOut = timelineRightBound

End Sub
'------------------------------------------------------------------------------
' FR: Calcule ou dessine une partie des liens de dependance visibles dans le GANTT.
' EN: Computes or draws part of the dependency links visible in GANTT.
'------------------------------------------------------------------------------
Private Sub FormatDependencyLine(ByVal shp As Shape, ByVal withArrow As Boolean)

    With shp.Line
        .ForeColor.RGB = RGB(120, 120, 120)
        .Weight = 1
        .DashStyle = msoLineSolid
        If withArrow Then
            .EndArrowheadStyle = msoArrowheadTriangle
        End If
    End With

End Sub
'------------------------------------------------------------------------------
' FR: Calcule ou dessine une partie des liens de dependance visibles dans le GANTT.
' EN: Computes or draws part of the dependency links visible in GANTT.
'------------------------------------------------------------------------------
Private Sub DrawLinkSegment( _
    ByVal ws As Worksheet, _
    ByVal shapeName As String, _
    ByVal x1 As Double, _
    ByVal y1 As Double, _
    ByVal x2 As Double, _
    ByVal y2 As Double, _
    ByVal withArrow As Boolean)

    Dim perfScope As clsPerfScope

    Dim shp As Shape

    Set perfScope = Profiler_BeginScope("DrawLinkSegment", "Shape Create")

    If Abs(x2 - x1) < 0.1 And Abs(y2 - y1) < 0.1 Then Exit Sub

    Set shp = ws.Shapes.AddLine(x1, y1, x2, y2)
    shp.Name = shapeName
    ApplyGanttRenderLinePlacement shp
    FormatDependencyLine shp, withArrow

End Sub

'------------------------------------------------------------------------------
' FR:
' Recharge le cache global des liens GANTT depuis tbl_LOGIC_LINKS avant dessin des dependances.
'
' EN:
' Reloads the global GANTT link cache from tbl_LOGIC_LINKS before dependency rendering.
'
' Entrees / Inputs:
' - tbl_LOGIC_LINKS via BuildExpandedLinksCacheFromLogicLinksTable.
'
' Sorties / Outputs:
' - gExpandedLinks remplace par le cache courant.
'
' Appele par / Called by:
' - DrawDependencyLinks.
'
' Notes:
' - Couplage direct CALC -> renderer; candidat a extraction Bridge/LinkProvider.
'------------------------------------------------------------------------------
Private Sub EnsureExpandedLinksCacheFromCalc()

    Set gExpandedLinks = Nothing
    Set gExpandedLinks = BuildExpandedLinksCacheFromLogicLinksTable()

End Sub
'------------------------------------------------------------------------------
' FR: Calcule ou dessine une partie des liens de dependance visibles dans le GANTT.
' EN: Computes or draws part of the dependency links visible in GANTT.
'------------------------------------------------------------------------------
Private Function BuildExpandedLinksCacheFromLogicLinksTable() As Object

    Dim perfScope As clsPerfScope
    Dim network As clsParsedPlanningNetwork
    Dim link As clsParsedPlanningLink
    Dim d As Object
    Dim linkCol As Collection
    Dim tokenInfo As Object
    Dim r As Long
    Dim succId As String
    Dim predId As String
    Dim linkType As String

    Set perfScope = Profiler_BeginScope("BuildExpandedLinksCacheFromLogicLinksTable", "Excel Read")
    Set d = CreateObject("Scripting.Dictionary")

    On Error GoTo SafeExit

    Set network = ParsedPlanningNetwork_LoadCanonical()

    If Not network.HasColumn("Succ ID") Then GoTo SafeExit
    If Not network.HasColumn("Pred ID") Then GoTo SafeExit
    If Not network.HasColumn("Link Type") Then GoTo SafeExit
    If Not network.HasColumn("Lag") Then GoTo SafeExit

    For r = 1 To network.Count

        Set link = network.Item(r)
        succId = link.SuccId
        predId = link.PredId
        linkType = link.LinkType

        If succId = "" Then GoTo NextRow
        If predId = "" Then GoTo NextRow
        If linkType <> "FS" And linkType <> "SS" And linkType <> "FF" Then GoTo NextRow

        If Not d.Exists(succId) Then
            Set linkCol = New Collection
            d.Add succId, linkCol
        End If

        Set tokenInfo = CreateObject("Scripting.Dictionary")
        tokenInfo("PredID") = predId
        tokenInfo("LinkType") = linkType
        tokenInfo("Lag") = link.Lag
        tokenInfo("RawToken") = link.RawToken

        d(succId).Add tokenInfo

NextRow:
    Next r

SafeExit:
    Set BuildExpandedLinksCacheFromLogicLinksTable = d

End Function
'------------------------------------------------------------------------------
' FR: Calcule ou dessine une partie des liens de dependance visibles dans le GANTT.
' EN: Computes or draws part of the dependency links visible in GANTT.
'------------------------------------------------------------------------------
Private Function HasExpandedLinksAvailable() As Boolean

    On Error GoTo SafeExit

    If gExpandedLinks Is Nothing Then Exit Function
    If gExpandedLinks.Count <= 0 Then Exit Function

    HasExpandedLinksAvailable = True
    Exit Function

SafeExit:
    HasExpandedLinksAvailable = False

End Function
'------------------------------------------------------------------------------
' FR: Calcule ou dessine une partie des liens de dependance visibles dans le GANTT.
' EN: Computes or draws part of the dependency links visible in GANTT.
'------------------------------------------------------------------------------
Private Sub GetLinkAnchorTypes( _
    ByVal linkType As String, _
    ByRef predAnchorType As String, _
    ByRef succAnchorType As String)

    Select Case UCase$(Trim$(linkType))
        Case "SS"
            predAnchorType = LINK_ANCHOR_START
            succAnchorType = LINK_ANCHOR_START

        Case "FF"
            predAnchorType = LINK_ANCHOR_FINISH
            succAnchorType = LINK_ANCHOR_FINISH

        Case Else
            predAnchorType = LINK_ANCHOR_FINISH
            succAnchorType = LINK_ANCHOR_START
    End Select

End Sub
'------------------------------------------------------------------------------
' FR: Calcule ou dessine une partie des liens de dependance visibles dans le GANTT.
' EN: Computes or draws part of the dependency links visible in GANTT.
'------------------------------------------------------------------------------
Private Function GetLinkReferenceDate( _
    ByVal taskId As String, _
    ByVal anchorType As String, _
    ByVal baseById As Object, _
    ByVal testById As Object, _
    ByVal isTestMode As Boolean) As Variant

    Select Case UCase$(Trim$(anchorType))
        Case LINK_ANCHOR_START
            GetLinkReferenceDate = GanttLive_GetDisplayStart(taskId, baseById, testById, isTestMode)

        Case Else
            GetLinkReferenceDate = GanttLive_GetDisplayFinish(taskId, baseById, testById, isTestMode)
    End Select

End Function
'------------------------------------------------------------------------------
' FR: Execute le helper Get Task Anchor Point dans le workflow de rendu GANTT.
' EN: Runs the Get Task Anchor Point helper in the GANTT rendering workflow.
'------------------------------------------------------------------------------
Private Sub GetTaskAnchorPoint( _
    ByVal ws As Worksheet, _
    ByVal mapWBS As Object, _
    ByRef dataArr As Variant, _
    ByVal hasChildren As Object, _
    ByVal projectStart As Variant, _
    ByVal totalDays As Long, _
    ByVal dataRow As Long, _
    ByVal isFinishSide As Boolean, _
    ByRef xOut As Double, _
    ByRef yOut As Double, _
    ByVal baseById As Object, _
    ByVal testById As Object, _
    ByVal isTestMode As Boolean)

    Dim ganttRow As Long
    Dim wbs As String
    Dim idVal As String
    Dim startVal As Variant
    Dim finishVal As Variant
    Dim durationVal As Double
    Dim sizeVal As Double

    Dim timelineLeftBound As Double
    Dim timelineRightBound As Double

    ganttRow = FIRST_TASK_ROW + dataRow - 1
    wbs = NormalizeWBS(CStr(dataArr(dataRow, mapWBS("WBS"))))
    idVal = Trim$(CStr(dataArr(dataRow, mapWBS("ID"))))

    startVal = GetRenderStartForCurrentScale(GanttLive_GetDisplayStart(idVal, baseById, testById, isTestMode))
    finishVal = GetRenderFinishForCurrentScale(GanttLive_GetDisplayFinish(idVal, baseById, testById, isTestMode))

    If Not HasValue(startVal) Or Not HasValue(finishVal) Then Exit Sub

    durationVal = CDbl(finishVal) - CDbl(startVal) + 1

    timelineLeftBound = ws.cells(HEADER_ROW_2, FIRST_TIMELINE_COL).Left + LINK_EDGE_PADDING
    timelineRightBound = ws.cells(HEADER_ROW_2, FIRST_TIMELINE_COL + totalDays - 1).Left + _
                         ws.cells(HEADER_ROW_2, FIRST_TIMELINE_COL + totalDays - 1).Width - LINK_EDGE_PADDING

    yOut = ws.cells(ganttRow, FIRST_TIMELINE_COL).Top + (ws.rows(ganttRow).Height / 2)

    If hasChildren.Exists(wbs) Then
        If isFinishSide Then
            xOut = TimelineRightAfterFinish(ws, projectStart, finishVal)
        Else
            xOut = TimelineLeft(ws, projectStart, startVal)
        End If

    ElseIf durationVal <= 1 Then
        sizeVal = ws.rows(ganttRow).Height - 6

        If isFinishSide Then
            xOut = GetTaskMidX(ws, projectStart, startVal) + (sizeVal / 2)
        Else
            xOut = GetTaskMidX(ws, projectStart, startVal) - (sizeVal / 2)
        End If

    Else
        If isFinishSide Then
            xOut = TimelineRightAfterFinish(ws, projectStart, finishVal)
        Else
            xOut = TimelineLeft(ws, projectStart, startVal)
        End If
    End If

    If xOut < timelineLeftBound Then xOut = timelineLeftBound
    If xOut > timelineRightBound Then xOut = timelineRightBound

End Sub
'------------------------------------------------------------------------------
' FR: Execute le helper Get Task Anchor Point By Type dans le workflow de rendu GANTT.
' EN: Runs the Get Task Anchor Point By Type helper in the GANTT rendering workflow.
'------------------------------------------------------------------------------
Private Sub GetTaskAnchorPointByType( _
    ByVal ws As Worksheet, _
    ByVal mapWBS As Object, _
    ByRef dataArr As Variant, _
    ByVal hasChildren As Object, _
    ByVal projectStart As Variant, _
    ByVal totalDays As Long, _
    ByVal dataRow As Long, _
    ByVal anchorType As String, _
    ByRef xOut As Double, _
    ByRef yOut As Double, _
    ByVal baseById As Object, _
    ByVal testById As Object, _
    ByVal isTestMode As Boolean)

    GetTaskAnchorPointBySide ws, mapWBS, dataArr, hasChildren, projectStart, totalDays, dataRow, _
                             IIf(UCase$(Trim$(anchorType)) = LINK_ANCHOR_START, "LEFT", "RIGHT"), _
                             xOut, yOut, baseById, testById, isTestMode

End Sub
'------------------------------------------------------------------------------
' FR: Execute le helper Get Task Finish Entry Point dans le workflow de rendu GANTT.
' EN: Runs the Get Task Finish Entry Point helper in the GANTT rendering workflow.
'------------------------------------------------------------------------------
Private Sub GetTaskFinishEntryPoint( _
    ByVal ws As Worksheet, _
    ByVal mapWBS As Object, _
    ByRef dataArr As Variant, _
    ByVal projectStart As Variant, _
    ByVal totalDays As Long, _
    ByVal dataRow As Long, _
    ByRef xOut As Double, _
    ByRef yOut As Double, _
    ByVal baseById As Object, _
    ByVal testById As Object, _
    ByVal isTestMode As Boolean)

    Dim ganttRow As Long
    Dim idVal As String
    Dim startVal As Variant
    Dim finishVal As Variant
    Dim durationVal As Double
    Dim sizeVal As Double
    Dim timelineLeftBound As Double
    Dim timelineRightBound As Double

    ganttRow = FIRST_TASK_ROW + dataRow - 1
    idVal = Trim$(CStr(dataArr(dataRow, mapWBS("ID"))))

    startVal = GetRenderStartForCurrentScale(GanttLive_GetDisplayStart(idVal, baseById, testById, isTestMode))
    finishVal = GetRenderFinishForCurrentScale(GanttLive_GetDisplayFinish(idVal, baseById, testById, isTestMode))

    If Not HasValue(startVal) Or Not HasValue(finishVal) Then Exit Sub

    durationVal = CDbl(finishVal) - CDbl(startVal) + 1

    timelineLeftBound = ws.cells(HEADER_ROW_2, FIRST_TIMELINE_COL).Left + LINK_EDGE_PADDING
    timelineRightBound = ws.cells(HEADER_ROW_2, FIRST_TIMELINE_COL + totalDays - 1).Left + _
                         ws.cells(HEADER_ROW_2, FIRST_TIMELINE_COL + totalDays - 1).Width - LINK_EDGE_PADDING

    yOut = ws.cells(ganttRow, FIRST_TIMELINE_COL).Top + (ws.rows(ganttRow).Height / 2)

    If durationVal <= 1 Then
        sizeVal = ws.rows(ganttRow).Height - 6
        xOut = GetTaskMidX(ws, projectStart, startVal) + (sizeVal / 2)
    Else
        xOut = TimelineRightAfterFinish(ws, projectStart, finishVal)
    End If

    If xOut < timelineLeftBound Then xOut = timelineLeftBound
    If xOut > timelineRightBound Then xOut = timelineRightBound

End Sub
'------------------------------------------------------------------------------
' FR: Execute le helper Get Task Left X dans le workflow de rendu GANTT.
' EN: Runs the Get Task Left X helper in the GANTT rendering workflow.
'------------------------------------------------------------------------------
Private Function GetTaskLeftX( _
    ByVal ws As Worksheet, _
    ByVal projectStart As Variant, _
    ByVal taskDate As Variant) As Double

    GetTaskLeftX = TimelineLeft(ws, projectStart, taskDate)

End Function
'------------------------------------------------------------------------------
' FR: Execute le helper Get Task Right X dans le workflow de rendu GANTT.
' EN: Runs the Get Task Right X helper in the GANTT rendering workflow.
'------------------------------------------------------------------------------
Private Function GetTaskRightX( _
    ByVal ws As Worksheet, _
    ByVal projectStart As Variant, _
    ByVal taskDate As Variant) As Double

    GetTaskRightX = TimelineRightAfterFinish(ws, projectStart, taskDate)

End Function
'------------------------------------------------------------------------------
' FR: Execute le helper Get Task Mid X dans le workflow de rendu GANTT.
' EN: Runs the Get Task Mid X helper in the GANTT rendering workflow.
'------------------------------------------------------------------------------
Public Function GetTaskMidX( _
    ByVal ws As Worksheet, _
    ByVal projectStart As Variant, _
    ByVal taskDate As Variant) As Double

    GetTaskMidX = TimelineDateRangeMidX(ws, projectStart, taskDate, taskDate)

End Function
'------------------------------------------------------------------------------
' FR: Execute le helper Get Task Anchor Point By Side dans le workflow de rendu GANTT.
' EN: Runs the Get Task Anchor Point By Side helper in the GANTT rendering workflow.
'------------------------------------------------------------------------------
Private Sub GetTaskAnchorPointBySide( _
    ByVal ws As Worksheet, _
    ByVal mapWBS As Object, _
    ByRef dataArr As Variant, _
    ByVal hasChildren As Object, _
    ByVal projectStart As Variant, _
    ByVal totalDays As Long, _
    ByVal dataRow As Long, _
    ByVal anchorSide As String, _
    ByRef xOut As Double, _
    ByRef yOut As Double, _
    ByVal baseById As Object, _
    ByVal testById As Object, _
    ByVal isTestMode As Boolean)

    Dim ganttRow As Long
    Dim wbs As String
    Dim idVal As String
    Dim startVal As Variant
    Dim finishVal As Variant
    Dim durationVal As Double
    Dim sizeVal As Double
    Dim timelineLeftBound As Double
    Dim timelineRightBound As Double

    ganttRow = FIRST_TASK_ROW + dataRow - 1
    wbs = NormalizeWBS(CStr(dataArr(dataRow, mapWBS("WBS"))))
    idVal = Trim$(CStr(dataArr(dataRow, mapWBS("ID"))))

    startVal = GetRenderStartForCurrentScale(GanttLive_GetDisplayStart(idVal, baseById, testById, isTestMode))
    finishVal = GetRenderFinishForCurrentScale(GanttLive_GetDisplayFinish(idVal, baseById, testById, isTestMode))

    If Not HasValue(startVal) Or Not HasValue(finishVal) Then Exit Sub

    durationVal = CDbl(finishVal) - CDbl(startVal) + 1

    timelineLeftBound = ws.cells(HEADER_ROW_2, FIRST_TIMELINE_COL).Left + LINK_EDGE_PADDING
    timelineRightBound = ws.cells(HEADER_ROW_2, FIRST_TIMELINE_COL + totalDays - 1).Left + _
                         ws.cells(HEADER_ROW_2, FIRST_TIMELINE_COL + totalDays - 1).Width - LINK_EDGE_PADDING

    yOut = ws.cells(ganttRow, FIRST_TIMELINE_COL).Top + (ws.rows(ganttRow).Height / 2)

    If hasChildren.Exists(wbs) Then
        Select Case UCase$(Trim$(anchorSide))
            Case "LEFT"
                xOut = TimelineLeft(ws, projectStart, startVal)
            Case "RIGHT"
                xOut = TimelineRightAfterFinish(ws, projectStart, finishVal)
            Case Else
                xOut = GetTaskMidX(ws, projectStart, startVal)
        End Select

    ElseIf durationVal <= 1 Then
        sizeVal = ws.rows(ganttRow).Height - 6

        Select Case UCase$(Trim$(anchorSide))
            Case "LEFT"
                xOut = GetTaskMidX(ws, projectStart, startVal) - (sizeVal / 2)
            Case "RIGHT"
                xOut = GetTaskMidX(ws, projectStart, startVal) + (sizeVal / 2)
            Case Else
                xOut = GetTaskMidX(ws, projectStart, startVal)
        End Select

    Else
        Select Case UCase$(Trim$(anchorSide))
            Case "LEFT"
                xOut = TimelineLeft(ws, projectStart, startVal)
            Case "RIGHT"
                xOut = TimelineRightAfterFinish(ws, projectStart, finishVal)
            Case Else
                xOut = GetTaskMidX(ws, projectStart, startVal)
        End Select
    End If

    If xOut < timelineLeftBound Then xOut = timelineLeftBound
    If xOut > timelineRightBound Then xOut = timelineRightBound

End Sub
