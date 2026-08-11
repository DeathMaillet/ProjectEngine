Attribute VB_Name = "mod_GanttVisualRegression"
Option Explicit

'===============================================================================
' MODULE : mod_GanttVisualRegression
' DOMAINE / DOMAIN : Gantt
'
' FR
' Harnais de preuve du contrat Gantt Visual Regression sur des copies de test.
' N'appartient a aucun workflow produit et ne doit pas etre appele en usage normal.
'
' EN
' Proof harness for the Gantt Visual Regression contract on test copies.
' Is not production workflow code and must not run during normal use.
'
' CONTRATS / CONTRACTS : GanttVR_CaptureSignature, GanttVR_CaptureToFile, GanttVR_SaveSignature, GanttVR_CompareSignatureFiles, GanttVR_CompareToReport, GanttVR_SmokeSelfCompare, GanttVR_SmokeDetectIntentionalChange
' CALLBACKS EXTERNES / EXTERNAL CALLBACKS : Aucun / None
'===============================================================================

Private Const GANTT_VR_VERSION As String = "1.0"
Private Const GANTT_VR_SHEET As String = "GANTT"
Private Const GANTT_VR_FIRST_TIMELINE_COL As Long = 11
Private Const GANTT_VR_FIRST_TASK_ROW As Long = 5
Private Const GANTT_VR_HEADER_ROW_1 As Long = 3
Private Const GANTT_VR_HEADER_ROW_2 As Long = 4
Private Const GANTT_VR_GEOMETRY_TOLERANCE As Double = 0.01
Private Const GANTT_VR_MAX_CELL_SCAN_ROWS As Long = 2000
Private Const GANTT_VR_MAX_CELL_SCAN_COLS As Long = 1000
Private Const GANTT_VR_TRACE_CAPTURE As Boolean = False

'------------------------------------------------------------------------------
' FR: Capture l'etat visuel courant du GANTT dans un dictionnaire de signature.
' EN: Captures the current visual state of GANTT into a signature dictionary.
'------------------------------------------------------------------------------
Public Function GanttVR_CaptureSignature(Optional ByVal sheetName As String = GANTT_VR_SHEET) As Object

    Dim signature As Object
    Dim ws As Worksheet

    Set signature = CreateObject("Scripting.Dictionary")

    On Error GoTo SafeExit

    Set ws = ThisWorkbook.Worksheets(sheetName)

    GanttVR_DebugTrace "meta:start"
    GanttVR_AddMetaSignature signature, ws
    GanttVR_DebugTrace "sheet:start"
    GanttVR_AddSheetSignature signature, ws
    GanttVR_DebugTrace "shapes:start"
    GanttVR_AddShapeSignatures signature, ws
    GanttVR_DebugTrace "timeline:start"
    GanttVR_AddTimelineSignature signature, ws
    GanttVR_DebugTrace "leftpane:start"
    GanttVR_AddLeftPaneSignature signature, ws
    GanttVR_DebugTrace "merges:start"
    GanttVR_AddMergedCellsSignature signature, ws
    GanttVR_DebugTrace "capture:done"

SafeExit:
    Set GanttVR_CaptureSignature = signature

End Function

'------------------------------------------------------------------------------
' FR: Capture le GANTT et sauvegarde la signature dans un fichier TSV.
' EN: Captures GANTT and saves the signature to a TSV file.
'------------------------------------------------------------------------------
Public Function GanttVR_CaptureToFile(Optional ByVal outputPath As String = "") As String

    Dim signature As Object
    Dim finalPath As String

    If Trim$(outputPath) = "" Then
        finalPath = GanttVR_DefaultSignaturePath("gantt_signature")
    Else
        finalPath = outputPath
    End If

    Set signature = GanttVR_CaptureSignature(GANTT_VR_SHEET)
    GanttVR_SaveSignature signature, finalPath

    GanttVR_CaptureToFile = finalPath

End Function

'------------------------------------------------------------------------------
' FR: Sauvegarde un dictionnaire de signature sous forme TSV deterministe.
' EN: Saves a signature dictionary as deterministic TSV.
'------------------------------------------------------------------------------
Public Sub GanttVR_SaveSignature(ByVal signature As Object, ByVal outputPath As String)

    Dim fso As Object
    Dim stream As Object
    Dim keys As Object
    Dim key As Variant

    If signature Is Nothing Then Exit Sub
    If Trim$(outputPath) = "" Then Exit Sub

    Set fso = CreateObject("Scripting.FileSystemObject")
    GanttVR_EnsureFolder fso.GetParentFolderName(outputPath)

    Set keys = GanttVR_SortedKeys(signature)
    Set stream = fso.CreateTextFile(outputPath, True, False)

    stream.WriteLine "#GANTT_VISUAL_SIGNATURE" & vbTab & GANTT_VR_VERSION
    For Each key In keys
        stream.WriteLine GanttVR_EncodeField(CStr(key)) & vbTab & GanttVR_EncodeField(CStr(signature(CStr(key))))
    Next key

    stream.Close

End Sub

'------------------------------------------------------------------------------
' FR: Compare deux fichiers de signature et retourne le nombre de differences.
' EN: Compares two signature files and returns the number of differences.
'------------------------------------------------------------------------------
Public Function GanttVR_CompareSignatureFiles( _
    ByVal beforePath As String, _
    ByVal afterPath As String, _
    Optional ByVal reportPath As String = "") As Long

    Dim beforeSig As Object
    Dim afterSig As Object
    Dim reportText As String
    Dim diffCount As Long

    Set beforeSig = GanttVR_LoadSignature(beforePath)
    Set afterSig = GanttVR_LoadSignature(afterPath)

    reportText = GanttVR_BuildComparisonReport(beforeSig, afterSig, beforePath, afterPath, diffCount)

    If Trim$(reportPath) <> "" Then
        GanttVR_WriteTextFile reportPath, reportText
    End If

    GanttVR_CompareSignatureFiles = diffCount

End Function

'------------------------------------------------------------------------------
' FR: Compare deux signatures et sauvegarde un rapport Markdown.
' EN: Compares two signatures and saves a Markdown report.
'------------------------------------------------------------------------------
Public Function GanttVR_CompareToReport( _
    ByVal beforePath As String, _
    ByVal afterPath As String, _
    Optional ByVal reportPath As String = "") As String

    Dim finalReportPath As String
    Dim diffCount As Long

    If Trim$(reportPath) = "" Then
        finalReportPath = GanttVR_DefaultSignaturePath("gantt_visual_compare", ".md")
    Else
        finalReportPath = reportPath
    End If

    diffCount = GanttVR_CompareSignatureFiles(beforePath, afterPath, finalReportPath)
    GanttVR_CompareToReport = finalReportPath

End Function

'------------------------------------------------------------------------------
' FR: Smoke interne: deux captures successives doivent etre identiques.
' EN: Internal smoke: two consecutive captures must be identical.
'------------------------------------------------------------------------------
Public Function GanttVR_SmokeSelfCompare(Optional ByVal reportPath As String = "") As Boolean

    Dim beforePath As String
    Dim afterPath As String
    Dim finalReportPath As String
    Dim diffCount As Long

    beforePath = GanttVR_DefaultSignaturePath("gantt_self_before")
    afterPath = GanttVR_DefaultSignaturePath("gantt_self_after")

    GanttVR_CaptureToFile beforePath
    GanttVR_CaptureToFile afterPath

    If Trim$(reportPath) = "" Then
        finalReportPath = GanttVR_DefaultSignaturePath("gantt_self_compare", ".md")
    Else
        finalReportPath = reportPath
    End If

    diffCount = GanttVR_CompareSignatureFiles(beforePath, afterPath, finalReportPath)
    GanttVR_SmokeSelfCompare = (diffCount = 0)

End Function

'------------------------------------------------------------------------------
' FR: Smoke interne: verifie qu'une modification volontaire est detectee.
' EN: Internal smoke: verifies that an intentional modification is detected.
'------------------------------------------------------------------------------
Public Function GanttVR_SmokeDetectIntentionalChange(Optional ByVal reportPath As String = "") As Boolean

    Dim ws As Worksheet
    Dim beforePath As String
    Dim afterPath As String
    Dim finalReportPath As String
    Dim shp As Shape
    Dim targetShape As Shape
    Dim oldLeft As Double
    Dim diffCount As Long

    On Error GoTo SafeExit

    Set ws = ThisWorkbook.Worksheets(GANTT_VR_SHEET)

    For Each shp In ws.Shapes
        If GanttVR_IsImportantShapeName(shp.Name) Then
            Set targetShape = shp
            Exit For
        End If
    Next shp

    If targetShape Is Nothing Then Exit Function

    beforePath = GanttVR_DefaultSignaturePath("gantt_mutation_before")
    afterPath = GanttVR_DefaultSignaturePath("gantt_mutation_after")

    GanttVR_CaptureToFile beforePath

    oldLeft = targetShape.Left
    targetShape.Left = oldLeft + 1#
    GanttVR_CaptureToFile afterPath
    targetShape.Left = oldLeft

    If Trim$(reportPath) = "" Then
        finalReportPath = GanttVR_DefaultSignaturePath("gantt_mutation_compare", ".md")
    Else
        finalReportPath = reportPath
    End If

    diffCount = GanttVR_CompareSignatureFiles(beforePath, afterPath, finalReportPath)
    GanttVR_SmokeDetectIntentionalChange = (diffCount > 0)
    Exit Function

SafeExit:
    On Error Resume Next
    If Not targetShape Is Nothing Then targetShape.Left = oldLeft
    GanttVR_SmokeDetectIntentionalChange = False

End Function

'------------------------------------------------------------------------------
' FR: Ajoute la map Meta Signature a la structure cible fournie par l'appelant.
' EN: Adds the Meta Signature map to the target structure supplied by the caller.
'------------------------------------------------------------------------------

Private Sub GanttVR_AddMetaSignature(ByVal signature As Object, ByVal ws As Worksheet)

    GanttVR_Add signature, "META|HarnessVersion", GANTT_VR_VERSION
    GanttVR_Add signature, "META|WorkbookName", ThisWorkbook.Name
    GanttVR_Add signature, "META|SheetName", ws.Name
    GanttVR_Add signature, "META|ShapeCount", CStr(ws.Shapes.Count)
    GanttVR_Add signature, "META|UsedRangeAddress", ws.UsedRange.Address(False, False)

    On Error Resume Next
    GanttVR_Add signature, "STATE|GanttViewMode", GetGanttViewMode()
    GanttVR_Add signature, "STATE|TimelineScaleMode", GetGanttTimelineScaleMode()
    GanttVR_Add signature, "STATE|ShowCriticalPath", CStr(GetGanttShowCriticalPath())
    GanttVR_Add signature, "STATE|PendingRenderMode", GanttLive_GetPendingRenderMode()
    GanttVR_Add signature, "STATE|ActiveSimulationMode", GanttLive_GetActiveSimulationMode()
    GanttVR_Add signature, "STATE|IsScenarioActive", CStr(GanttLive_IsScenarioActive())
    GanttVR_Add signature, "STATE|IsLiveTestActive", CStr(GanttLive_IsLiveTestActive())
    GanttVR_Add signature, "STATE|HasPendingGeometryRepair", CStr(Gantt_HasPendingGeometryRepair())
    GanttVR_Add signature, "STATE|PendingGeometryRepairCount", CStr(Gantt_PendingGeometryRepairCount())
    GanttVR_Add signature, "STATE|DragIsWatching", CStr(GanttDrag_IsWatching())
    On Error GoTo 0

End Sub

'------------------------------------------------------------------------------
' FR: Ajoute la map Sheet Signature a la structure cible fournie par l'appelant.
' EN: Adds the Sheet Signature map to the target structure supplied by the caller.
'------------------------------------------------------------------------------

Private Sub GanttVR_AddSheetSignature(ByVal signature As Object, ByVal ws As Worksheet)

    Dim lastRow As Long
    Dim lastCol As Long
    Dim r As Long
    Dim c As Long

    lastRow = GanttVR_LastGanttRow(ws)
    lastCol = GanttVR_LastTimelineCol(ws)

    If lastRow < 1 Then lastRow = 1
    If lastCol < 1 Then lastCol = 1
    If lastRow > GANTT_VR_MAX_CELL_SCAN_ROWS Then lastRow = GANTT_VR_MAX_CELL_SCAN_ROWS
    If lastCol > GANTT_VR_MAX_CELL_SCAN_COLS Then lastCol = GANTT_VR_MAX_CELL_SCAN_COLS

    GanttVR_Add signature, "SHEET|LastRow", CStr(lastRow)
    GanttVR_Add signature, "SHEET|LastCol", CStr(lastCol)
    On Error Resume Next
    GanttVR_Add signature, "SHEET|FreezePanes", CStr(ActiveWindow.FreezePanes)
    On Error GoTo 0

    For c = 1 To lastCol
        GanttVR_Add signature, "COL|" & GanttVR_PadLong(c) & "|Width", GanttVR_FormatDouble(ws.Columns(c).ColumnWidth)
        GanttVR_Add signature, "COL|" & GanttVR_PadLong(c) & "|Hidden", CStr(ws.Columns(c).Hidden)
    Next c

    For r = 1 To lastRow
        GanttVR_Add signature, "ROW|" & GanttVR_PadLong(r) & "|Height", GanttVR_FormatDouble(ws.Rows(r).RowHeight)
        GanttVR_Add signature, "ROW|" & GanttVR_PadLong(r) & "|Hidden", CStr(ws.Rows(r).Hidden)
    Next r

End Sub

'------------------------------------------------------------------------------
' FR: Ajoute la map Shape Signatures a la structure cible fournie par l'appelant.
' EN: Adds the Shape Signatures map to the target structure supplied by the caller.
'------------------------------------------------------------------------------

Private Sub GanttVR_AddShapeSignatures(ByVal signature As Object, ByVal ws As Worksheet)

    Dim shp As Shape

    For Each shp In ws.Shapes
        GanttVR_DebugTrace "shape:" & shp.Name
        GanttVR_AddShapeSignature signature, shp
    Next shp

End Sub

'------------------------------------------------------------------------------
' FR: Traite la map Debug Trace sans modifier les donnees d'entree.
' EN: Handles the Debug Trace map without mutating input data.
'------------------------------------------------------------------------------

Private Sub GanttVR_DebugTrace(ByVal messageText As String)

    Dim fso As Object
    Dim stream As Object
    Dim folderPath As String
    Dim filePath As String

    If Not GANTT_VR_TRACE_CAPTURE Then Exit Sub

    On Error Resume Next
    folderPath = ThisWorkbook.Path & Application.PathSeparator & "performance_audit" & Application.PathSeparator & "gantt_visual_regression"
    GanttVR_EnsureFolder folderPath
    filePath = folderPath & Application.PathSeparator & "_last_capture_trace.txt"
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set stream = fso.OpenTextFile(filePath, 8, True, False)
    stream.WriteLine Format$(Now, "yyyy-mm-dd hh:nn:ss") & vbTab & messageText
    stream.Close
    On Error GoTo 0

End Sub

'------------------------------------------------------------------------------
' FR: Ajoute la map Shape Signature a la structure cible fournie par l'appelant.
' EN: Adds the Shape Signature map to the target structure supplied by the caller.
'------------------------------------------------------------------------------

Private Sub GanttVR_AddShapeSignature(ByVal signature As Object, ByVal shp As Shape)

    Dim prefix As String

    prefix = "SHAPE|" & GanttVR_NormalizeKeyPart(shp.Name) & "|"

    GanttVR_Add signature, prefix & "Name", shp.Name
    GanttVR_Add signature, prefix & "Family", GanttVR_ShapeFamily(shp.Name)
    GanttVR_Add signature, prefix & "Type", GanttVR_SafeLongProperty(shp, "Type")
    GanttVR_Add signature, prefix & "AutoShapeType", GanttVR_SafeLongProperty(shp, "AutoShapeType")
    GanttVR_Add signature, prefix & "Left", GanttVR_FormatDouble(shp.Left)
    GanttVR_Add signature, prefix & "Top", GanttVR_FormatDouble(shp.Top)
    GanttVR_Add signature, prefix & "Width", GanttVR_FormatDouble(shp.Width)
    GanttVR_Add signature, prefix & "Height", GanttVR_FormatDouble(shp.Height)
    GanttVR_Add signature, prefix & "Rotation", GanttVR_FormatDouble(shp.Rotation)
    GanttVR_Add signature, prefix & "Visible", CStr(shp.Visible)
    GanttVR_Add signature, prefix & "ZOrderPosition", GanttVR_SafeLongProperty(shp, "ZOrderPosition")
    GanttVR_Add signature, prefix & "OnAction", GanttVR_SafeStringProperty(shp, "OnAction")
    GanttVR_Add signature, prefix & "Placement", GanttVR_SafeLongProperty(shp, "Placement")
    GanttVR_Add signature, prefix & "Locked", GanttVR_SafeBoolProperty(shp, "Locked")
    GanttVR_Add signature, prefix & "AlternativeText", GanttVR_SafeStringProperty(shp, "AlternativeText")
    GanttVR_Add signature, prefix & "NameHash", CStr(Len(shp.Name)) & ":" & shp.Name

    If shp.Type <> msoLine Then GanttVR_AddFillSignature signature, prefix, shp
    GanttVR_AddLineSignature signature, prefix, shp
    If GanttVR_ShouldCaptureShapeText(shp.Name) Then GanttVR_AddTextSignature signature, prefix, shp
    If shp.Type = msoGroup Then GanttVR_AddGroupSignature signature, prefix, shp

End Sub

'------------------------------------------------------------------------------
' FR: Ajoute la map Fill Signature a la structure cible fournie par l'appelant.
' EN: Adds the Fill Signature map to the target structure supplied by the caller.
'------------------------------------------------------------------------------

Private Sub GanttVR_AddFillSignature(ByVal signature As Object, ByVal prefix As String, ByVal shp As Shape)

    On Error Resume Next
    GanttVR_Add signature, prefix & "Fill.Visible", CStr(shp.Fill.Visible)
    GanttVR_Add signature, prefix & "Fill.ForeColor.RGB", CStr(shp.Fill.ForeColor.RGB)
    GanttVR_Add signature, prefix & "Fill.BackColor.RGB", CStr(shp.Fill.BackColor.RGB)
    GanttVR_Add signature, prefix & "Fill.Transparency", GanttVR_FormatDouble(shp.Fill.Transparency)
    GanttVR_Add signature, prefix & "Fill.Type", CStr(shp.Fill.Type)
    GanttVR_Add signature, prefix & "Fill.Pattern", CStr(shp.Fill.Pattern)
    On Error GoTo 0

End Sub

'------------------------------------------------------------------------------
' FR: Ajoute la map Line Signature a la structure cible fournie par l'appelant.
' EN: Adds the Line Signature map to the target structure supplied by the caller.
'------------------------------------------------------------------------------

Private Sub GanttVR_AddLineSignature(ByVal signature As Object, ByVal prefix As String, ByVal shp As Shape)

    On Error Resume Next
    GanttVR_Add signature, prefix & "Line.Visible", CStr(shp.Line.Visible)
    GanttVR_Add signature, prefix & "Line.ForeColor.RGB", CStr(shp.Line.ForeColor.RGB)
    GanttVR_Add signature, prefix & "Line.BackColor.RGB", CStr(shp.Line.BackColor.RGB)
    GanttVR_Add signature, prefix & "Line.Transparency", GanttVR_FormatDouble(shp.Line.Transparency)
    GanttVR_Add signature, prefix & "Line.Weight", GanttVR_FormatDouble(shp.Line.Weight)
    GanttVR_Add signature, prefix & "Line.DashStyle", CStr(shp.Line.DashStyle)
    GanttVR_Add signature, prefix & "Line.Style", CStr(shp.Line.Style)
    If shp.Type = msoLine Then
        GanttVR_Add signature, prefix & "Line.BeginArrowheadStyle", CStr(shp.Line.BeginArrowheadStyle)
        GanttVR_Add signature, prefix & "Line.EndArrowheadStyle", CStr(shp.Line.EndArrowheadStyle)
        GanttVR_Add signature, prefix & "Line.BeginArrowheadLength", CStr(shp.Line.BeginArrowheadLength)
        GanttVR_Add signature, prefix & "Line.EndArrowheadLength", CStr(shp.Line.EndArrowheadLength)
        GanttVR_Add signature, prefix & "Line.BeginArrowheadWidth", CStr(shp.Line.BeginArrowheadWidth)
        GanttVR_Add signature, prefix & "Line.EndArrowheadWidth", CStr(shp.Line.EndArrowheadWidth)
    End If
    On Error GoTo 0

End Sub

'------------------------------------------------------------------------------
' FR: Ajoute la map Text Signature a la structure cible fournie par l'appelant.
' EN: Adds the Text Signature map to the target structure supplied by the caller.
'------------------------------------------------------------------------------

Private Sub GanttVR_AddTextSignature(ByVal signature As Object, ByVal prefix As String, ByVal shp As Shape)

    Dim txt As String

    txt = ""
    On Error Resume Next
    If shp.TextFrame2.HasText Then txt = shp.TextFrame2.TextRange.Text
    If Err.Number <> 0 Then Err.Clear
    If txt = "" Then txt = shp.TextFrame.Characters.Text

    GanttVR_Add signature, prefix & "Text", txt
    GanttVR_Add signature, prefix & "Text.Length", CStr(Len(txt))
    If Len(txt) > 0 Then
        GanttVR_Add signature, prefix & "TextFrame2.VerticalAnchor", CStr(shp.TextFrame2.VerticalAnchor)
        GanttVR_Add signature, prefix & "TextFrame2.MarginLeft", GanttVR_FormatDouble(shp.TextFrame2.MarginLeft)
        GanttVR_Add signature, prefix & "TextFrame2.MarginRight", GanttVR_FormatDouble(shp.TextFrame2.MarginRight)
        GanttVR_Add signature, prefix & "TextFrame2.MarginTop", GanttVR_FormatDouble(shp.TextFrame2.MarginTop)
        GanttVR_Add signature, prefix & "TextFrame2.MarginBottom", GanttVR_FormatDouble(shp.TextFrame2.MarginBottom)
        GanttVR_Add signature, prefix & "Font.Name", shp.TextFrame2.TextRange.Font.Name
        GanttVR_Add signature, prefix & "Font.Size", GanttVR_FormatDouble(shp.TextFrame2.TextRange.Font.Size)
        GanttVR_Add signature, prefix & "Font.Bold", CStr(shp.TextFrame2.TextRange.Font.Bold)
        GanttVR_Add signature, prefix & "Font.Italic", CStr(shp.TextFrame2.TextRange.Font.Italic)
        GanttVR_Add signature, prefix & "Font.FillColor", CStr(shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB)
    End If
    On Error GoTo 0

End Sub

'------------------------------------------------------------------------------
' FR: Ajoute la map Group Signature a la structure cible fournie par l'appelant.
' EN: Adds the Group Signature map to the target structure supplied by the caller.
'------------------------------------------------------------------------------

Private Sub GanttVR_AddGroupSignature(ByVal signature As Object, ByVal prefix As String, ByVal shp As Shape)

    Dim i As Long
    Dim names As String

    On Error Resume Next
    GanttVR_Add signature, prefix & "GroupItems.Count", CStr(shp.GroupItems.Count)
    For i = 1 To shp.GroupItems.Count
        If names <> "" Then names = names & "|"
        names = names & shp.GroupItems(i).Name
    Next i
    GanttVR_Add signature, prefix & "GroupItems.Names", names
    On Error GoTo 0

End Sub

'------------------------------------------------------------------------------
' FR: Ajoute la map Timeline Signature a la structure cible fournie par l'appelant.
' EN: Adds the Timeline Signature map to the target structure supplied by the caller.
'------------------------------------------------------------------------------

Private Sub GanttVR_AddTimelineSignature(ByVal signature As Object, ByVal ws As Worksheet)

    Dim lastCol As Long
    Dim c As Long
    Dim keyCol As String

    lastCol = GanttVR_LastTimelineCol(ws)
    If lastCol < GANTT_VR_FIRST_TIMELINE_COL Then Exit Sub

    GanttVR_Add signature, "TIMELINE|FirstCol", CStr(GANTT_VR_FIRST_TIMELINE_COL)
    GanttVR_Add signature, "TIMELINE|LastCol", CStr(lastCol)
    GanttVR_Add signature, "TIMELINE|SlotCount", CStr(lastCol - GANTT_VR_FIRST_TIMELINE_COL + 1)

    For c = GANTT_VR_FIRST_TIMELINE_COL To lastCol
        keyCol = GanttVR_PadLong(c)
        GanttVR_AddCellSignature signature, ws.Cells(GANTT_VR_HEADER_ROW_1, c), "TIMELINE|COL|" & keyCol & "|Header1"
        GanttVR_AddCellSignature signature, ws.Cells(GANTT_VR_HEADER_ROW_2, c), "TIMELINE|COL|" & keyCol & "|Header2"
    Next c

End Sub

'------------------------------------------------------------------------------
' FR: Ajoute la map Left Pane Signature a la structure cible fournie par l'appelant.
' EN: Adds the Left Pane Signature map to the target structure supplied by the caller.
'------------------------------------------------------------------------------

Private Sub GanttVR_AddLeftPaneSignature(ByVal signature As Object, ByVal ws As Worksheet)

    Dim lastRow As Long
    Dim r As Long
    Dim c As Long

    lastRow = GanttVR_LastGanttRow(ws)
    If lastRow < 1 Then Exit Sub

    For r = 1 To lastRow
        For c = 1 To 10
            GanttVR_AddCellSignature signature, ws.Cells(r, c), "CELL|" & GanttVR_PadLong(r) & "|" & GanttVR_PadLong(c)
        Next c
    Next r

End Sub

'------------------------------------------------------------------------------
' FR: Ajoute la map Merged Cells Signature a la structure cible fournie par l'appelant.
' EN: Adds the Merged Cells Signature map to the target structure supplied by the caller.
'------------------------------------------------------------------------------

Private Sub GanttVR_AddMergedCellsSignature(ByVal signature As Object, ByVal ws As Worksheet)

    Dim seen As Object
    Dim lastRow As Long
    Dim lastCol As Long

    Set seen = CreateObject("Scripting.Dictionary")

    lastRow = GanttVR_LastGanttRow(ws)
    lastCol = GanttVR_LastTimelineCol(ws)
    If lastRow < GANTT_VR_HEADER_ROW_2 Then lastRow = GANTT_VR_HEADER_ROW_2
    If lastCol < 10 Then lastCol = 10
    If lastRow > GANTT_VR_MAX_CELL_SCAN_ROWS Then lastRow = GANTT_VR_MAX_CELL_SCAN_ROWS
    If lastCol > GANTT_VR_MAX_CELL_SCAN_COLS Then lastCol = GANTT_VR_MAX_CELL_SCAN_COLS

    GanttVR_AddMergedCellsInRange signature, seen, ws.Range(ws.Cells(1, 1), ws.Cells(GANTT_VR_HEADER_ROW_2, lastCol))
    GanttVR_AddMergedCellsInRange signature, seen, ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, 10))

End Sub

'------------------------------------------------------------------------------
' FR: Ajoute la map Merged Cells In Range a la structure cible fournie par l'appelant.
' EN: Adds the Merged Cells In Range map to the target structure supplied by the caller.
'------------------------------------------------------------------------------

Private Sub GanttVR_AddMergedCellsInRange(ByVal signature As Object, ByVal seen As Object, ByVal scanRange As Range)

    Dim cell As Range
    Dim areaAddress As String

    For Each cell In scanRange.Cells
        If cell.MergeCells Then
            areaAddress = cell.MergeArea.Address(False, False)
            If Not seen.Exists(areaAddress) Then
                seen(areaAddress) = True
                GanttVR_Add signature, "MERGE|" & areaAddress, areaAddress
            End If
        End If
    Next cell

End Sub

'------------------------------------------------------------------------------
' FR: Ajoute la map Cell Signature a la structure cible fournie par l'appelant.
' EN: Adds the Cell Signature map to the target structure supplied by the caller.
'------------------------------------------------------------------------------

Private Sub GanttVR_AddCellSignature(ByVal signature As Object, ByVal cell As Range, ByVal keyPrefix As String)

    On Error Resume Next
    GanttVR_Add signature, keyPrefix & "|Address", cell.Address(False, False)
    GanttVR_Add signature, keyPrefix & "|Value", GanttVR_CellText(cell)
    GanttVR_Add signature, keyPrefix & "|NumberFormat", CStr(cell.NumberFormat)
    GanttVR_Add signature, keyPrefix & "|Interior.Color", CStr(cell.Interior.Color)
    GanttVR_Add signature, keyPrefix & "|Font.Name", CStr(cell.Font.Name)
    GanttVR_Add signature, keyPrefix & "|Font.Size", GanttVR_FormatDouble(cell.Font.Size)
    GanttVR_Add signature, keyPrefix & "|Font.Bold", CStr(cell.Font.Bold)
    GanttVR_Add signature, keyPrefix & "|Font.Italic", CStr(cell.Font.Italic)
    GanttVR_Add signature, keyPrefix & "|Font.Color", CStr(cell.Font.Color)
    GanttVR_Add signature, keyPrefix & "|HorizontalAlignment", CStr(cell.HorizontalAlignment)
    GanttVR_Add signature, keyPrefix & "|VerticalAlignment", CStr(cell.VerticalAlignment)
    If cell.MergeCells Then
        GanttVR_Add signature, keyPrefix & "|MergeArea", cell.MergeArea.Address(False, False)
    Else
        GanttVR_Add signature, keyPrefix & "|MergeArea", ""
    End If
    GanttVR_Add signature, keyPrefix & "|Locked", CStr(cell.Locked)
    On Error GoTo 0

End Sub

'------------------------------------------------------------------------------
' FR: Charge la map Signature depuis sa source proprietaire sans appliquer de politique aval.
' EN: Loads the Signature map from its owning source without applying downstream policy.
'------------------------------------------------------------------------------

Private Function GanttVR_LoadSignature(ByVal inputPath As String) As Object

    Dim fso As Object
    Dim stream As Object
    Dim d As Object
    Dim lineText As String
    Dim parts As Variant

    Set d = CreateObject("Scripting.Dictionary")
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FileExists(inputPath) Then
        Set GanttVR_LoadSignature = d
        Exit Function
    End If

    Set stream = fso.OpenTextFile(inputPath, 1, False)
    Do While Not stream.AtEndOfStream
        lineText = stream.ReadLine
        If Left$(lineText, 1) <> "#" Then
            parts = Split(lineText, vbTab)
            If UBound(parts) >= 1 Then
                d(GanttVR_DecodeField(CStr(parts(0)))) = GanttVR_DecodeField(CStr(parts(1)))
            End If
        End If
    Loop
    stream.Close

    Set GanttVR_LoadSignature = d

End Function

'------------------------------------------------------------------------------
' FR: Indique si la valeur Capture Shape Text satisfait la condition attendue, sans modifier les donnees source.
' EN: Returns whether the Capture Shape Text value satisfies the expected condition without mutating source data.
'------------------------------------------------------------------------------

Private Function GanttVR_ShouldCaptureShapeText(ByVal shapeName As String) As Boolean

    Select Case GanttVR_ShapeFamily(shapeName)
        Case "DEPENDENCY", "CONSTRAINT", "TODAY_LINE"
            GanttVR_ShouldCaptureShapeText = False
        Case Else
            GanttVR_ShouldCaptureShapeText = True
    End Select

End Function

'------------------------------------------------------------------------------
' FR: Construit la map Comparison Report a partir des donnees fournies par l'appelant.
' EN: Builds the Comparison Report map from data supplied by the caller.
'------------------------------------------------------------------------------

Private Function GanttVR_BuildComparisonReport( _
    ByVal beforeSig As Object, _
    ByVal afterSig As Object, _
    ByVal beforePath As String, _
    ByVal afterPath As String, _
    ByRef diffCount As Long) As String

    Dim report As String
    Dim allKeys As Object
    Dim key As Variant
    Dim beforeValue As String
    Dim afterValue As String
    Dim addedCount As Long
    Dim removedCount As Long
    Dim changedCount As Long
    Dim identicalCount As Long
    Dim details As String
    Dim maxDetails As Long

    Set allKeys = GanttVR_UnionSortedKeys(beforeSig, afterSig)
    maxDetails = 500

    For Each key In allKeys
        If GanttVR_ShouldIgnoreComparisonKey(CStr(key)) Then GoTo NextComparisonKey

        If beforeSig.Exists(CStr(key)) And afterSig.Exists(CStr(key)) Then
            beforeValue = CStr(beforeSig(CStr(key)))
            afterValue = CStr(afterSig(CStr(key)))
            If GanttVR_ValuesEquivalent(CStr(key), beforeValue, afterValue) Then
                identicalCount = identicalCount + 1
            Else
                changedCount = changedCount + 1
                diffCount = diffCount + 1
                If changedCount <= maxDetails Then
                    details = details & GanttVR_DiffBlock("DIFFERENCE", CStr(key), beforeValue, afterValue)
                End If
            End If
        ElseIf beforeSig.Exists(CStr(key)) Then
            removedCount = removedCount + 1
            diffCount = diffCount + 1
            If removedCount <= maxDetails Then
                details = details & GanttVR_DiffBlock("SUPPRIME / REMOVED", CStr(key), CStr(beforeSig(CStr(key))), "")
            End If
        Else
            addedCount = addedCount + 1
            diffCount = diffCount + 1
            If addedCount <= maxDetails Then
                details = details & GanttVR_DiffBlock("AJOUTE / ADDED", CStr(key), "", CStr(afterSig(CStr(key))))
            End If
        End If

NextComparisonKey:
    Next key

    report = "# Gantt Visual Regression Comparison" & vbCrLf & vbCrLf
    report = report & "- Harness version: `" & GANTT_VR_VERSION & "`" & vbCrLf
    report = report & "- Before: `" & beforePath & "`" & vbCrLf
    report = report & "- After: `" & afterPath & "`" & vbCrLf & vbCrLf
    report = report & "## Summary" & vbCrLf & vbCrLf
    report = report & "| Metric | Count |" & vbCrLf
    report = report & "|---|---:|" & vbCrLf
    report = report & "| Identical properties | " & CStr(identicalCount) & " |" & vbCrLf
    report = report & "| Added properties | " & CStr(addedCount) & " |" & vbCrLf
    report = report & "| Removed properties | " & CStr(removedCount) & " |" & vbCrLf
    report = report & "| Changed properties | " & CStr(changedCount) & " |" & vbCrLf
    report = report & "| Total differences | " & CStr(diffCount) & " |" & vbCrLf & vbCrLf

    If diffCount = 0 Then
        report = report & "## Result" & vbCrLf & vbCrLf & "No visual signature differences detected." & vbCrLf
    Else
        report = report & "## Differences" & vbCrLf & vbCrLf & details
        If diffCount > maxDetails Then
            report = report & vbCrLf & "_Report truncated to " & CStr(maxDetails) & " detail blocks._" & vbCrLf
        End If
    End If

    GanttVR_BuildComparisonReport = report

End Function

'------------------------------------------------------------------------------
' FR: Indique si la valeur Ignore Comparison Key satisfait la condition attendue, sans modifier les donnees source.
' EN: Returns whether the Ignore Comparison Key value satisfies the expected condition without mutating source data.
'------------------------------------------------------------------------------

Private Function GanttVR_ShouldIgnoreComparisonKey(ByVal keyText As String) As Boolean

    Select Case keyText
        Case "META|WorkbookName"
            GanttVR_ShouldIgnoreComparisonKey = True
    End Select

End Function

'------------------------------------------------------------------------------
' FR: Retourne la valeur Diff Block sans modifier les donnees d'entree.
' EN: Returns the Diff Block value without mutating input data.
'------------------------------------------------------------------------------

Private Function GanttVR_DiffBlock(ByVal statusText As String, ByVal keyText As String, ByVal beforeValue As String, ByVal afterValue As String) As String

    Dim block As String
    Dim entityName As String
    Dim propName As String

    entityName = GanttVR_EntityFromKey(keyText)
    propName = GanttVR_PropertyFromKey(keyText)

    block = "### " & entityName & vbCrLf & vbCrLf
    block = block & "- Property: `" & propName & "`" & vbCrLf
    block = block & "- Status: **" & statusText & "**" & vbCrLf
    block = block & "- Before: `" & GanttVR_MarkdownInline(beforeValue) & "`" & vbCrLf
    block = block & "- After: `" & GanttVR_MarkdownInline(afterValue) & "`" & vbCrLf & vbCrLf

    GanttVR_DiffBlock = block

End Function

'------------------------------------------------------------------------------
' FR: Retourne la valeur Entity From Key sans modifier les donnees d'entree.
' EN: Returns the Entity From Key value without mutating input data.
'------------------------------------------------------------------------------

Private Function GanttVR_EntityFromKey(ByVal keyText As String) As String

    Dim parts As Variant

    parts = Split(keyText, "|")
    If UBound(parts) >= 1 Then
        If CStr(parts(0)) = "SHAPE" Then
            GanttVR_EntityFromKey = CStr(parts(1))
        ElseIf UBound(parts) >= 2 Then
            GanttVR_EntityFromKey = CStr(parts(0)) & " " & CStr(parts(1)) & " " & CStr(parts(2))
        Else
            GanttVR_EntityFromKey = CStr(parts(0)) & " " & CStr(parts(1))
        End If
    Else
        GanttVR_EntityFromKey = keyText
    End If

End Function

'------------------------------------------------------------------------------
' FR: Retourne la valeur Property From Key sans modifier les donnees d'entree.
' EN: Returns the Property From Key value without mutating input data.
'------------------------------------------------------------------------------

Private Function GanttVR_PropertyFromKey(ByVal keyText As String) As String

    Dim pos As Long

    pos = InStrRev(keyText, "|")
    If pos > 0 Then
        GanttVR_PropertyFromKey = Mid$(keyText, pos + 1)
    Else
        GanttVR_PropertyFromKey = keyText
    End If

End Function

'------------------------------------------------------------------------------
' FR: Retourne la valeur Values Equivalent sans modifier les donnees d'entree.
' EN: Returns the Values Equivalent value without mutating input data.
'------------------------------------------------------------------------------

Private Function GanttVR_ValuesEquivalent(ByVal keyText As String, ByVal beforeValue As String, ByVal afterValue As String) As Boolean

    If beforeValue = afterValue Then
        GanttVR_ValuesEquivalent = True
        Exit Function
    End If

    If GanttVR_IsGeometryKey(keyText) Then
        If IsNumeric(beforeValue) And IsNumeric(afterValue) Then
            GanttVR_ValuesEquivalent = (Abs(CDbl(beforeValue) - CDbl(afterValue)) <= GANTT_VR_GEOMETRY_TOLERANCE)
            Exit Function
        End If
    End If

    GanttVR_ValuesEquivalent = False

End Function

'------------------------------------------------------------------------------
' FR: Indique si la valeur Geometry Key satisfait la condition attendue, sans modifier les donnees source.
' EN: Returns whether the Geometry Key value satisfies the expected condition without mutating source data.
'------------------------------------------------------------------------------

Private Function GanttVR_IsGeometryKey(ByVal keyText As String) As Boolean

    Dim propName As String

    propName = GanttVR_PropertyFromKey(keyText)

    Select Case propName
        Case "Left", "Top", "Width", "Height", "Rotation", "Line.Weight", "Fill.Transparency", "Line.Transparency", "Font.Size"
            GanttVR_IsGeometryKey = True
    End Select

End Function

'------------------------------------------------------------------------------
' FR: Ajoute la map Add a la structure cible fournie par l'appelant.
' EN: Adds the Add map to the target structure supplied by the caller.
'------------------------------------------------------------------------------

Private Sub GanttVR_Add(ByVal signature As Object, ByVal keyText As String, ByVal valueText As String)

    signature(keyText) = valueText

End Sub

'------------------------------------------------------------------------------
' FR: Retourne la reference Cell Text sans modifier les donnees d'entree.
' EN: Returns the Cell Text reference without mutating input data.
'------------------------------------------------------------------------------

Private Function GanttVR_CellText(ByVal cell As Range) As String

    On Error Resume Next
    If IsError(cell.Value2) Then
        GanttVR_CellText = "#ERROR"
    ElseIf IsEmpty(cell.Value2) Then
        GanttVR_CellText = ""
    Else
        GanttVR_CellText = CStr(cell.Value2)
    End If
    On Error GoTo 0

End Function

'------------------------------------------------------------------------------
' FR: Retourne la valeur Shape Family sans modifier les donnees d'entree.
' EN: Returns the Shape Family value without mutating input data.
'------------------------------------------------------------------------------

Private Function GanttVR_ShapeFamily(ByVal shapeName As String) As String

    If shapeName = "TODAY_LINE" Then
        GanttVR_ShapeFamily = "TODAY_LINE"
    ElseIf Left$(shapeName, 5) = "TASK_" Then
        GanttVR_ShapeFamily = "TASK"
    ElseIf Left$(shapeName, 3) = "MS_" Then
        GanttVR_ShapeFamily = "MILESTONE"
    ElseIf Left$(shapeName, 4) = "SUM_" Then
        GanttVR_ShapeFamily = "SUMMARY"
    ElseIf Left$(shapeName, 4) = "DEP_" Then
        GanttVR_ShapeFamily = "DEPENDENCY"
    ElseIf Left$(shapeName, 5) = "CSTR_" Then
        GanttVR_ShapeFamily = "CONSTRAINT"
    ElseIf InStr(1, shapeName, "Gantt", vbTextCompare) > 0 Or InStr(1, shapeName, "GANTT", vbTextCompare) > 0 Then
        GanttVR_ShapeFamily = "GANTT_UI"
    Else
        GanttVR_ShapeFamily = "OTHER"
    End If

End Function

'------------------------------------------------------------------------------
' FR: Indique si la valeur Important Shape Name satisfait la condition attendue, sans modifier les donnees source.
' EN: Returns whether the Important Shape Name value satisfies the expected condition without mutating source data.
'------------------------------------------------------------------------------

Private Function GanttVR_IsImportantShapeName(ByVal shapeName As String) As Boolean

    Select Case True
        Case shapeName = "TODAY_LINE"
            GanttVR_IsImportantShapeName = True
        Case Left$(shapeName, 5) = "TASK_"
            GanttVR_IsImportantShapeName = True
        Case Left$(shapeName, 3) = "MS_"
            GanttVR_IsImportantShapeName = True
        Case Left$(shapeName, 4) = "SUM_"
            GanttVR_IsImportantShapeName = True
        Case Left$(shapeName, 4) = "DEP_"
            GanttVR_IsImportantShapeName = True
        Case Left$(shapeName, 5) = "CSTR_"
            GanttVR_IsImportantShapeName = True
    End Select

End Function

'------------------------------------------------------------------------------
' FR: Retourne la reference Last Timeline Col sans modifier les donnees d'entree.
' EN: Returns the Last Timeline Col reference without mutating input data.
'------------------------------------------------------------------------------

Private Function GanttVR_LastTimelineCol(ByVal ws As Worksheet) As Long

    Dim c1 As Long
    Dim c2 As Long

    c1 = ws.Cells(GANTT_VR_HEADER_ROW_1, ws.Columns.Count).End(xlToLeft).Column
    c2 = ws.Cells(GANTT_VR_HEADER_ROW_2, ws.Columns.Count).End(xlToLeft).Column

    If c1 > c2 Then
        GanttVR_LastTimelineCol = c1
    Else
        GanttVR_LastTimelineCol = c2
    End If

End Function

'------------------------------------------------------------------------------
' FR: Retourne la reference Last Gantt Row sans modifier les donnees d'entree.
' EN: Returns the Last Gantt Row reference without mutating input data.
'------------------------------------------------------------------------------

Private Function GanttVR_LastGanttRow(ByVal ws As Worksheet) As Long

    Dim lastRow As Long

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < GANTT_VR_FIRST_TASK_ROW Then lastRow = ws.UsedRange.Rows.Count
    If lastRow < 1 Then lastRow = 1

    GanttVR_LastGanttRow = lastRow

End Function

'------------------------------------------------------------------------------
' FR: Retourne la reference Last Used Row sans modifier les donnees d'entree.
' EN: Returns the Last Used Row reference without mutating input data.
'------------------------------------------------------------------------------

Private Function GanttVR_LastUsedRow(ByVal ws As Worksheet) As Long

    On Error Resume Next
    GanttVR_LastUsedRow = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    If GanttVR_LastUsedRow < 1 Then GanttVR_LastUsedRow = 1
    On Error GoTo 0

End Function

'------------------------------------------------------------------------------
' FR: Retourne la reference Last Used Col sans modifier les donnees d'entree.
' EN: Returns the Last Used Col reference without mutating input data.
'------------------------------------------------------------------------------

Private Function GanttVR_LastUsedCol(ByVal ws As Worksheet) As Long

    On Error Resume Next
    GanttVR_LastUsedCol = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
    If GanttVR_LastUsedCol < 1 Then GanttVR_LastUsedCol = 1
    On Error GoTo 0

End Function

'------------------------------------------------------------------------------
' FR: Retourne la valeur Default Signature Path sans modifier les donnees d'entree.
' EN: Returns the Default Signature Path value without mutating input data.
'------------------------------------------------------------------------------

Private Function GanttVR_DefaultSignaturePath(ByVal prefixText As String, Optional ByVal extensionText As String = ".tsv") As String

    Dim folderPath As String

    folderPath = ThisWorkbook.Path & Application.PathSeparator & "performance_audit" & Application.PathSeparator & "gantt_visual_regression"
    GanttVR_EnsureFolder folderPath
    GanttVR_DefaultSignaturePath = folderPath & Application.PathSeparator & prefixText & "_" & Format$(Now, "yyyymmdd_hhnnss") & extensionText

End Function

'------------------------------------------------------------------------------
' FR: Cree ou remet en conformite la map Folder de maniere idempotente.
' EN: Creates or restores the Folder map to a compliant state idempotently.
'------------------------------------------------------------------------------

Private Sub GanttVR_EnsureFolder(ByVal folderPath As String)

    Dim fso As Object
    Dim parentPath As String

    If Trim$(folderPath) = "" Then Exit Sub

    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FolderExists(folderPath) Then Exit Sub

    parentPath = fso.GetParentFolderName(folderPath)
    If Trim$(parentPath) <> "" Then
        If Not fso.FolderExists(parentPath) Then GanttVR_EnsureFolder parentPath
    End If

    fso.CreateFolder folderPath

End Sub

'------------------------------------------------------------------------------
' FR: Ecrit ou synchronise Write Text File dans le stockage possede par le domaine.
' EN: Writes or synchronizes Write Text File in the store owned by the domain.
'------------------------------------------------------------------------------

Private Sub GanttVR_WriteTextFile(ByVal filePath As String, ByVal textValue As String)

    Dim fso As Object
    Dim stream As Object

    Set fso = CreateObject("Scripting.FileSystemObject")
    GanttVR_EnsureFolder fso.GetParentFolderName(filePath)

    Set stream = fso.CreateTextFile(filePath, True, False)
    stream.Write textValue
    stream.Close

End Sub

'------------------------------------------------------------------------------
' FR: Trie la collection Sorted Keys en place selon l'ordre deterministic attendu par le consommateur.
' EN: Sorts the Sorted Keys collection in place using the deterministic order expected by the consumer.
'------------------------------------------------------------------------------

Private Function GanttVR_SortedKeys(ByVal d As Object) As Object

    Dim keys As Collection
    Dim key As Variant
    Dim values() As String
    Dim i As Long

    Set keys = New Collection
    If d Is Nothing Then
        Set GanttVR_SortedKeys = keys
        Exit Function
    End If
    If d.Count = 0 Then
        Set GanttVR_SortedKeys = keys
        Exit Function
    End If

    ReDim values(1 To d.Count)
    i = 0
    For Each key In d.Keys
        i = i + 1
        values(i) = CStr(key)
    Next key

    If UBound(values) > LBound(values) Then GanttVR_QuickSortStrings values, LBound(values), UBound(values)

    For i = LBound(values) To UBound(values)
        keys.Add values(i)
    Next i

    Set GanttVR_SortedKeys = keys

End Function

'------------------------------------------------------------------------------
' FR: Trie la valeur Quick Sort Strings en place selon l'ordre deterministic attendu par le consommateur.
' EN: Sorts the Quick Sort Strings value in place using the deterministic order expected by the consumer.
'------------------------------------------------------------------------------

Private Sub GanttVR_QuickSortStrings(ByRef values() As String, ByVal firstIndex As Long, ByVal lastIndex As Long)

    Dim lowIndex As Long
    Dim highIndex As Long
    Dim pivotValue As String
    Dim tmpValue As String

    lowIndex = firstIndex
    highIndex = lastIndex
    pivotValue = values((firstIndex + lastIndex) \ 2)

    Do While lowIndex <= highIndex
        Do While values(lowIndex) < pivotValue
            lowIndex = lowIndex + 1
        Loop
        Do While values(highIndex) > pivotValue
            highIndex = highIndex - 1
        Loop
        If lowIndex <= highIndex Then
            tmpValue = values(lowIndex)
            values(lowIndex) = values(highIndex)
            values(highIndex) = tmpValue
            lowIndex = lowIndex + 1
            highIndex = highIndex - 1
        End If
    Loop

    If firstIndex < highIndex Then GanttVR_QuickSortStrings values, firstIndex, highIndex
    If lowIndex < lastIndex Then GanttVR_QuickSortStrings values, lowIndex, lastIndex

End Sub

'------------------------------------------------------------------------------
' FR: Retourne la map Union Sorted Keys sans modifier les donnees d'entree.
' EN: Returns the Union Sorted Keys map without mutating input data.
'------------------------------------------------------------------------------

Private Function GanttVR_UnionSortedKeys(ByVal a As Object, ByVal b As Object) As Object

    Dim tmp As Object
    Dim key As Variant

    Set tmp = CreateObject("Scripting.Dictionary")

    If Not a Is Nothing Then
        For Each key In a.Keys
            tmp(CStr(key)) = True
        Next key
    End If

    If Not b Is Nothing Then
        For Each key In b.Keys
            tmp(CStr(key)) = True
        Next key
    End If

    Set GanttVR_UnionSortedKeys = GanttVR_SortedKeys(tmp)

End Function

'------------------------------------------------------------------------------
' FR: Retourne la valeur Encode Field sans modifier les donnees d'entree.
' EN: Returns the Encode Field value without mutating input data.
'------------------------------------------------------------------------------

Private Function GanttVR_EncodeField(ByVal valueText As String) As String

    valueText = Replace(valueText, "\", "\\")
    valueText = Replace(valueText, vbCrLf, "\n")
    valueText = Replace(valueText, vbCr, "\n")
    valueText = Replace(valueText, vbLf, "\n")
    valueText = Replace(valueText, vbTab, "\t")
    GanttVR_EncodeField = valueText

End Function

'------------------------------------------------------------------------------
' FR: Retourne la valeur Decode Field sans modifier les donnees d'entree.
' EN: Returns the Decode Field value without mutating input data.
'------------------------------------------------------------------------------

Private Function GanttVR_DecodeField(ByVal valueText As String) As String

    Dim result As String
    Dim i As Long
    Dim ch As String
    Dim nextCh As String

    result = ""
    i = 1
    Do While i <= Len(valueText)
        ch = Mid$(valueText, i, 1)
        If ch = "\" And i < Len(valueText) Then
            nextCh = Mid$(valueText, i + 1, 1)
            Select Case nextCh
                Case "n"
                    result = result & vbCrLf
                    i = i + 2
                Case "t"
                    result = result & vbTab
                    i = i + 2
                Case "\"
                    result = result & "\"
                    i = i + 2
                Case Else
                    result = result & ch
                    i = i + 1
            End Select
        Else
            result = result & ch
            i = i + 1
        End If
    Loop

    GanttVR_DecodeField = result

End Function

'------------------------------------------------------------------------------
' FR: Normalise ou formate Normalize Key Part selon le contrat canonique du composant.
' EN: Normalizes or formats Normalize Key Part according to the component contract.
'------------------------------------------------------------------------------

Private Function GanttVR_NormalizeKeyPart(ByVal valueText As String) As String

    valueText = Replace(valueText, "|", "_")
    valueText = Replace(valueText, vbTab, "_")
    valueText = Replace(valueText, vbCr, "_")
    valueText = Replace(valueText, vbLf, "_")
    GanttVR_NormalizeKeyPart = valueText

End Function

'------------------------------------------------------------------------------
' FR: Normalise ou formate Format Double selon le contrat canonique du composant.
' EN: Normalizes or formats Format Double according to the component contract.
'------------------------------------------------------------------------------

Private Function GanttVR_FormatDouble(ByVal value As Double) As String

    GanttVR_FormatDouble = Replace(Format$(value, "0.0000"), ",", ".")

End Function

'------------------------------------------------------------------------------
' FR: Retourne la valeur Pad Long sans modifier les donnees d'entree.
' EN: Returns the Pad Long value without mutating input data.
'------------------------------------------------------------------------------

Private Function GanttVR_PadLong(ByVal value As Long) As String

    GanttVR_PadLong = Right$("000000" & CStr(value), 6)

End Function

'------------------------------------------------------------------------------
' FR: Retourne la map Safe String Property sans modifier les donnees d'entree.
' EN: Returns the Safe String Property map without mutating input data.
'------------------------------------------------------------------------------

Private Function GanttVR_SafeStringProperty(ByVal obj As Object, ByVal propertyName As String) As String

    On Error GoTo SafeExit
    GanttVR_SafeStringProperty = CStr(CallByName(obj, propertyName, VbGet))
    Exit Function

SafeExit:
    GanttVR_SafeStringProperty = ""

End Function

'------------------------------------------------------------------------------
' FR: Retourne la map Safe Long Property sans modifier les donnees d'entree.
' EN: Returns the Safe Long Property map without mutating input data.
'------------------------------------------------------------------------------

Private Function GanttVR_SafeLongProperty(ByVal obj As Object, ByVal propertyName As String) As String

    On Error GoTo SafeExit
    GanttVR_SafeLongProperty = CStr(CLng(CallByName(obj, propertyName, VbGet)))
    Exit Function

SafeExit:
    GanttVR_SafeLongProperty = ""

End Function

'------------------------------------------------------------------------------
' FR: Retourne la map Safe Bool Property sans modifier les donnees d'entree.
' EN: Returns the Safe Bool Property map without mutating input data.
'------------------------------------------------------------------------------

Private Function GanttVR_SafeBoolProperty(ByVal obj As Object, ByVal propertyName As String) As String

    On Error GoTo SafeExit
    GanttVR_SafeBoolProperty = CStr(CBool(CallByName(obj, propertyName, VbGet)))
    Exit Function

SafeExit:
    GanttVR_SafeBoolProperty = ""

End Function

'------------------------------------------------------------------------------
' FR: Marque la valeur Markdown Inline dans la structure cible sans recalculer la condition source.
' EN: Marks the Markdown Inline value in the target structure without recalculating the source condition.
'------------------------------------------------------------------------------

Private Function GanttVR_MarkdownInline(ByVal valueText As String) As String

    valueText = Replace(valueText, "`", "'")
    valueText = Replace(valueText, vbCrLf, " ")
    valueText = Replace(valueText, vbCr, " ")
    valueText = Replace(valueText, vbLf, " ")
    If Len(valueText) > 300 Then valueText = Left$(valueText, 300) & "..."
    GanttVR_MarkdownInline = valueText

End Function
