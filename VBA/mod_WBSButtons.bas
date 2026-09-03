Attribute VB_Name = "mod_WBSButtons"
Option Explicit

'===============================================================================
' MODULE : mod_WBSButtons
' DOMAINE / DOMAIN : WBS
'
' FR
' Possede Reset/Armageddon, les boutons WBS, leur langue, le guide de saisie et la mise en forme des inputs.
' Ne doit pas contourner les contrats publics des autres domaines.
'
' EN
' Owns Reset/Armageddon, WBS buttons, their language, the onboarding guide and input formatting.
' Must not bypass public contracts owned by other domains.
'
' CONTRATS / CONTRACTS : Armageddon, Reset_Planning, WBS_EnsureOnboardingGuide, Ensure_WBS_Main_Buttons, WBS_ApplyLanguage, WBS_SetLanguage, WBS_CurrentLanguage
' CALLBACKS EXTERNES / EXTERNAL CALLBACKS : Aucun / None
'===============================================================================


Private Const RESET_PLANNING_EMPTY_WBS_ROWS As Long = 5

Private Const WBS_ONBOARDING_FONT_NAME As String = "Aptos Narrow"
Private Const WBS_ONBOARDING_FONT_SIZE As Double = 8
Private Const WBS_ONBOARDING_ROW_HEIGHT As Double = 14
Private Const WBS_QUICK_START_FR_LINE_1 As String = "Démarrage rapide : renseignez les colonnes portant le statut Requis. Les colonnes Optionnel peuvent rester vides et les colonnes Calculé sont renseignées automatiquement."
Private Const WBS_QUICK_START_FR_LINE_2 As String = "* Predecessors WBS, Baseline Start et Baseline Duration doivent ensemble fournir suffisamment d'informations pour positionner la tâche sur la chronologie."
Private Const WBS_QUICK_START_EN_LINE_1 As String = "Quick Start: Fill in the columns marked Required. Optional columns may be left blank, while Calculated columns are populated automatically."
Private Const WBS_QUICK_START_EN_LINE_2 As String = "* Predecessors WBS, Baseline Start and Baseline Duration must together provide enough information to position the task on the timeline."
Private gWBSLanguage As String
Private gWBSOnboardingLocalizedCheckCount As Long
Private gWBSOnboardingLocalizedRebuildCount As Long
Private gWBSOnboardingStructuralMutationCount As Long
Private gWBSOnboardingThreadedReadCount As Long
Private gWBSOnboardingThreadedAddCount As Long
Private gWBSOnboardingThreadedClearCount As Long
'------------------------------------------------------------------------------
' FR: Traite la reference Armageddon sans modifier les donnees d'entree.
' EN: Handles the Armageddon reference without mutating input data.
'------------------------------------------------------------------------------

Public Sub Armageddon(Optional ByVal skipConfirmation As Boolean = False)

    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim answer As VbMsgBoxResult
    Dim oldEvents As Boolean
    Dim oldScreenUpdating As Boolean

    oldEvents = Application.EnableEvents
    oldScreenUpdating = Application.ScreenUpdating

    On Error GoTo ErrHandler

    If Not skipConfirmation Then
        answer = MsgBox( _
            FormatPlanningConsoleMessageForCurrentLanguage(BiMsg( _
                "Cette action va vider le planning, le Dashboard, l'historique et les acquittements." & vbCrLf & _
                "Continuer ?", _
                "This will clear the planning, Dashboard, event history and acknowledgements." & vbCrLf & _
                "Continue?")), _
            vbQuestion + vbYesNo + vbDefaultButton2, _
            FormatPlanningConsoleMessageForCurrentLanguage(BiMsg("Full Reset", "Full Reset")))

        If answer <> vbYes Then Exit Sub
    End If

    Set ws = ThisWorkbook.Worksheets("WBS")
    Set tbl = ws.ListObjects("tbl_WBS")

    Application.EnableEvents = False
    Application.ScreenUpdating = False
    ResetPlanning_PrepareEmptyWBS ws, tbl
    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEvents

    Run_Full_Update
    Reset_Dashboard
    ClearPlanningWarningAcknowledgements
    ClearPlanningEventHistory
    Exit Sub

ErrHandler:
    On Error Resume Next
    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEvents
    On Error GoTo 0

    WBSButtons_ShowConsoleError _
        "Armageddon", _
        "Erreur dans Full Reset : " & Err.Description, _
        "Error in Full Reset: " & Err.Description

End Sub

'------------------------------------------------------------------------------
' FR: Reinitialise Reset Planning dans le perimetre possede par le composant.
' EN: Resets Reset Planning within the state owned by the component.
'------------------------------------------------------------------------------

Public Sub Reset_Planning()

    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim answer As VbMsgBoxResult
    Dim oldEvents As Boolean
    Dim oldScreenUpdating As Boolean

    On Error GoTo ErrHandler

    answer = MsgBox( _
        FormatPlanningConsoleMessageForCurrentLanguage(BiMsg( _
            "Cette action va vider le WBS et nettoyer les sorties calcul planning." & vbCrLf & _
            "Gantt, S-Curve, Dashboard, historique et acknowledgements ne seront pas modifies." & vbCrLf & vbCrLf & _
            "Continuer ?", _
            "This will clear the WBS and clean planning calculation outputs." & vbCrLf & _
            "Gantt, S-Curve, Dashboard, history and acknowledgements will not be modified." & vbCrLf & vbCrLf & _
            "Continue?")), _
        vbQuestion + vbYesNo + vbDefaultButton2, _
        FormatPlanningConsoleMessageForCurrentLanguage(BiMsg("Reset Planning", "Reset Planning")))

    If answer <> vbYes Then Exit Sub

    Set ws = ThisWorkbook.Worksheets("WBS")
    Set tbl = ws.ListObjects("tbl_WBS")

    oldEvents = Application.EnableEvents
    oldScreenUpdating = Application.ScreenUpdating
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    ResetPlanning_PrepareEmptyWBS ws, tbl

CleanReset:
    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEvents

    Run_Planning_Update
    Exit Sub

ErrHandler:
    On Error Resume Next
    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEvents
    On Error GoTo 0

    WBSButtons_ShowConsoleError _
        "Reset_Planning", _
        "Erreur dans Reset Planning : " & Err.Description, _
        "Error in Reset Planning: " & Err.Description

End Sub

'------------------------------------------------------------------------------
' FR: Reinitialise Reset Planning Prepare Empty WBS dans le perimetre possede par le composant.
' EN: Resets Reset Planning Prepare Empty WBS within the state owned by the component.
' FR - Effet de bord : efface uniquement les donnees ou objets cibles du contrat.
' EN - Side effect: clears only data or objects targeted by the contract.
'------------------------------------------------------------------------------

Private Sub ResetPlanning_PrepareEmptyWBS( _
    ByVal ws As Worksheet, _
    ByVal tbl As ListObject)

    If ws Is Nothing Then Exit Sub
    If tbl Is Nothing Then Exit Sub

    Do While tbl.ListRows.Count > RESET_PLANNING_EMPTY_WBS_ROWS
        tbl.ListRows(tbl.ListRows.Count).Delete
    Loop

    Do While tbl.ListRows.Count < RESET_PLANNING_EMPTY_WBS_ROWS
        tbl.ListRows.Add
    Loop

    If Not tbl.DataBodyRange Is Nothing Then
        tbl.DataBodyRange.ClearContents
        RestoreWBSFormulaColumns tbl
        ResetPlanning_ApplyWBSInputSetup tbl
    End If

End Sub

'------------------------------------------------------------------------------
' FR: Reinitialise Reset Planning Apply WBS Input Setup dans le perimetre possede par le composant.
' EN: Resets Reset Planning Apply WBS Input Setup within the state owned by the component.
'------------------------------------------------------------------------------

Private Sub ResetPlanning_ApplyWBSInputSetup(ByVal tbl As ListObject)

    If tbl Is Nothing Then Exit Sub
    If tbl.DataBodyRange Is Nothing Then Exit Sub

    ResetPlanning_ApplyListValidation tbl, VTS_COL_TASK_TYPE, "Task,Milestone,Level of Effort", _
        "Task Type", "Choose: Task, Milestone, or Level of Effort.", _
        "Invalid Task Type", "Allowed values: Task, Milestone, Level of Effort."

    ResetPlanning_ApplyListValidation tbl, VTS_COL_S, "Y,N", _
        "S", "Choose Y to show in Summary, N to hide.", _
        "Invalid S", "Allowed values: blank, Y, N."

    ResetPlanning_ApplyListValidation tbl, VTS_COL_CAL, CALENDAR_7D & "," & CALENDAR_6D & "," & CALENDAR_5D, _
        "Cal", "Choose: 7j/7, 6j/7, or 5j/7.", _
        "Invalid Cal", "Allowed values: blank, 7j/7, 6j/7, 5j/7."

    ResetPlanning_ApplyWBSFormats tbl

End Sub

'------------------------------------------------------------------------------
' FR: Reinitialise Reset Planning Apply List Validation dans le perimetre possede par le composant.
' EN: Resets Reset Planning Apply List Validation within the state owned by the component.
' FR - Effet de bord : efface uniquement les donnees ou objets cibles du contrat.
' EN - Side effect: clears only data or objects targeted by the contract.
'------------------------------------------------------------------------------

Private Sub ResetPlanning_ApplyListValidation( _
    ByVal tbl As ListObject, _
    ByVal columnKey As String, _
    ByVal listFormula As String, _
    ByVal inputTitle As String, _
    ByVal inputMessage As String, _
    ByVal errorTitle As String, _
    ByVal errorMessage As String)

    Dim rng As Range

    If Not WBS_TableHasSchemaColumn(tbl, columnKey) Then Exit Sub
    Set rng = SchemaListColumn(tbl, VTS_TABLE_WBS, columnKey).DataBodyRange
    If rng Is Nothing Then Exit Sub

    rng.NumberFormat = "@"
    With rng.Validation
        .Delete
        .Add Type:=xlValidateList, _
             AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, _
             Formula1:=listFormula
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = inputTitle
        .InputMessage = inputMessage
        .ErrorTitle = errorTitle
        .errorMessage = errorMessage
        .ShowInput = True
        .ShowError = True
    End With

End Sub

'------------------------------------------------------------------------------
' FR: Reinitialise Reset Planning Apply WBS Formats dans le perimetre possede par le composant.
' EN: Resets Reset Planning Apply WBS Formats within the state owned by the component.
'------------------------------------------------------------------------------

Private Sub ResetPlanning_ApplyWBSFormats(ByVal tbl As ListObject)

    Dim dateFormat As String

    dateFormat = Settings_GetDateNumberFormat()

    If WBS_TableHasSchemaColumn(tbl, VTS_COL_WBS) Then SchemaListColumn(tbl, VTS_TABLE_WBS, VTS_COL_WBS).DataBodyRange.NumberFormat = "@"
    If WBS_TableHasSchemaColumn(tbl, VTS_COL_ID) Then SchemaListColumn(tbl, VTS_TABLE_WBS, VTS_COL_ID).DataBodyRange.NumberFormat = "0"
    If WBS_TableHasSchemaColumn(tbl, VTS_COL_BASELINE_START) Then SchemaListColumn(tbl, VTS_TABLE_WBS, VTS_COL_BASELINE_START).DataBodyRange.NumberFormat = dateFormat
    If WBS_TableHasSchemaColumn(tbl, VTS_COL_BASELINE_FINISH) Then SchemaListColumn(tbl, VTS_TABLE_WBS, VTS_COL_BASELINE_FINISH).DataBodyRange.NumberFormat = dateFormat
    If WBS_TableHasSchemaColumn(tbl, VTS_COL_ACTUAL_START) Then SchemaListColumn(tbl, VTS_TABLE_WBS, VTS_COL_ACTUAL_START).DataBodyRange.NumberFormat = dateFormat
    If WBS_TableHasSchemaColumn(tbl, VTS_COL_ACTUAL_FINISH) Then SchemaListColumn(tbl, VTS_TABLE_WBS, VTS_COL_ACTUAL_FINISH).DataBodyRange.NumberFormat = dateFormat
    If WBS_TableHasSchemaColumn(tbl, VTS_COL_FORECAST_START) Then SchemaListColumn(tbl, VTS_TABLE_WBS, VTS_COL_FORECAST_START).DataBodyRange.NumberFormat = dateFormat
    If WBS_TableHasSchemaColumn(tbl, VTS_COL_FORECAST_FINISH) Then SchemaListColumn(tbl, VTS_TABLE_WBS, VTS_COL_FORECAST_FINISH).DataBodyRange.NumberFormat = dateFormat
    If WBS_TableHasSchemaColumn(tbl, VTS_COL_BASELINE_DURATION) Then SchemaListColumn(tbl, VTS_TABLE_WBS, VTS_COL_BASELINE_DURATION).DataBodyRange.NumberFormat = "0"
    If WBS_TableHasSchemaColumn(tbl, VTS_COL_ACTUAL_DURATION) Then SchemaListColumn(tbl, VTS_TABLE_WBS, VTS_COL_ACTUAL_DURATION).DataBodyRange.NumberFormat = "0"
    If WBS_TableHasSchemaColumn(tbl, VTS_COL_CALCULATED_DURATION) Then SchemaListColumn(tbl, VTS_TABLE_WBS, VTS_COL_CALCULATED_DURATION).DataBodyRange.NumberFormat = "0"

End Sub

'=====================================================
' mod_WBSButtons
'
' WBS main buttons + Task Type setup.
'
' Console routing:
' - confirmation utilisateur autorisee pour Reset Planning
' - les erreurs sont envoyées vers frmPlanningMessages
'=====================================================

'------------------------------------------------------------------------------
' FR: Garantit le layout d'aide WBS v1.0.1 et migre un ancien layout compatible.
' EN: Ensures the v1.0.1 WBS help layout and migrates a compatible legacy layout.
' FR - Effet de bord : insere la ligne 3 une seule fois si tbl_WBS commence en ligne 4.
' EN - Side effect: inserts row 3 exactly once when tbl_WBS starts on row 4.
'------------------------------------------------------------------------------
Public Sub WBS_EnsureOnboardingGuide()

    Dim ws As Worksheet

    Set ws = ThisWorkbook.Worksheets("WBS")
    Ensure_WBS_Onboarding_Guide ws

End Sub

'------------------------------------------------------------------------------
' FR: Migre et maintient le guide de saisie WBS sans dependre de la position des colonnes.
' EN: Migrates and maintains the WBS onboarding guide without relying on column positions.
' FR - Effet de bord : insere la ligne 3 une seule fois dans les anciens classeurs compatibles.
' EN - Side effect: inserts row 3 exactly once in compatible legacy workbooks.
'------------------------------------------------------------------------------
Private Sub Ensure_WBS_Onboarding_Guide( _
    ByVal ws As Worksheet, _
    Optional ByVal applyLocalizedContent As Boolean = True)

    Dim tbl As ListObject
    Dim statuses As Object
    Dim key As Variant
    Dim missingColumns As String
    Dim headerRow As Long
    Dim firstDataRow As Long
    Dim oldEvents As Boolean
    Dim oldScreenUpdating As Boolean
    Dim renamedLegacyProjectColumn As Boolean
    Dim structureNeedsMutation As Boolean
    Dim errorNumber As Long
    Dim errorSource As String
    Dim errorDescription As String

    If ws Is Nothing Then Exit Sub

    oldEvents = Application.EnableEvents
    oldScreenUpdating = Application.ScreenUpdating
    On Error GoTo ErrHandler

    Set tbl = ws.ListObjects("tbl_WBS")
    Set statuses = WBS_Onboarding_BuildStatusMap()

    If WBS_TableHasSchemaColumn(tbl, VTS_COL_PROJECT) And WBS_TableHasPhysicalColumn(tbl, "Package") Then
        Err.Raise vbObjectError + 2329, "Ensure_WBS_Onboarding_Guide", _
            "Ambiguous WBS schema: both Project and legacy Package columns are present."
    End If

    For Each key In statuses.Keys
        If Not WBS_TableHasSchemaColumn(tbl, CStr(key)) Then
            If CStr(key) <> VTS_COL_PROJECT Or Not WBS_TableHasPhysicalColumn(tbl, "Package") Then
                If missingColumns <> "" Then missingColumns = missingColumns & ", "
                missingColumns = missingColumns & CStr(key)
            End If
        End If
    Next key

    If missingColumns <> "" Then
        Err.Raise vbObjectError + 2330, "Ensure_WBS_Onboarding_Guide", _
            "Missing or renamed canonical WBS column(s): " & missingColumns
    End If

    headerRow = tbl.HeaderRowRange.Row
    If headerRow <= 1 Then
        Err.Raise vbObjectError + 2331, "Ensure_WBS_Onboarding_Guide", _
            "tbl_WBS has no row available immediately above its physical header for onboarding."
    End If

    structureNeedsMutation = Not WBS_TableHasSchemaColumn(tbl, VTS_COL_PROJECT)

    If Not structureNeedsMutation Then
        structureNeedsMutation = Not WBS_Onboarding_StructureIsCurrent(ws, tbl, statuses)
    End If

    If structureNeedsMutation Then
        Application.EnableEvents = False
        Application.ScreenUpdating = False

        If Not WBS_TableHasSchemaColumn(tbl, VTS_COL_PROJECT) Then
            tbl.ListColumns("Package").Name = SchemaCurrentColumnTitle(VTS_TABLE_WBS, VTS_COL_PROJECT)
            renamedLegacyProjectColumn = True
        End If

        Set tbl = ws.ListObjects("tbl_WBS")
        WBS_Onboarding_ApplyStructure ws, tbl, statuses
        gWBSOnboardingStructuralMutationCount = _
            gWBSOnboardingStructuralMutationCount + 1
    End If

    If Not tbl.DataBodyRange Is Nothing Then
        firstDataRow = tbl.DataBodyRange.Row
        If firstDataRow <> tbl.HeaderRowRange.Row + 1 Then
            Err.Raise vbObjectError + 2333, "Ensure_WBS_Onboarding_Guide", _
                "Unexpected first tbl_WBS data row. Expected the row immediately below the physical header, found " & _
                CStr(firstDataRow) & "."
        End If
    End If

    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEvents

    If applyLocalizedContent Then
        WBS_SetLanguage Settings_GetOwnerLanguage("WBS")
        WBS_Onboarding_ApplyLocalizedContent ws, tbl
    End If
    Exit Sub

ErrHandler:
    errorNumber = Err.Number
    errorSource = Err.Source
    errorDescription = Err.Description

    On Error Resume Next
    Application.EnableEvents = False
    If renamedLegacyProjectColumn Then
        Set tbl = ws.ListObjects("tbl_WBS")
        If WBS_TableHasSchemaColumn(tbl, VTS_COL_PROJECT) Then SchemaListColumn(tbl, VTS_TABLE_WBS, VTS_COL_PROJECT).Name = "Package"
    End If
    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEvents
    On Error GoTo 0

    Err.Raise errorNumber, errorSource, errorDescription

End Sub

'------------------------------------------------------------------------------
' FR: Resout la ligne d'onboarding relativement au vrai header de tbl_WBS.
' EN: Resolves the onboarding row relative to the physical tbl_WBS header.
'------------------------------------------------------------------------------
Private Function WBS_OnboardingHelpRow(ByVal tbl As ListObject) As Long

    If tbl Is Nothing Or tbl.HeaderRowRange.Row <= 1 Then
        Err.Raise vbObjectError + 2334, "WBS_OnboardingHelpRow", _
            "tbl_WBS has no valid onboarding row immediately above its physical header."
    End If
    WBS_OnboardingHelpRow = tbl.HeaderRowRange.Row - 1

End Function

Public Function WBS_OnboardingHarness_CheckStructure() As String
    Dim ws As Worksheet
    Dim tbl As ListObject

    Set ws = ThisWorkbook.Worksheets("WBS")
    Set tbl = ws.ListObjects("tbl_WBS")
    WBS_OnboardingHarness_CheckStructure = CStr( _
        WBS_Onboarding_StructureIsCurrent(ws, tbl, WBS_Onboarding_BuildStatusMap()))
End Function

Public Function WBS_OnboardingHarness_CheckLocalized( _
    Optional ByVal languageCode As String = "") As String
    Dim ws As Worksheet
    Dim tbl As ListObject

    Set ws = ThisWorkbook.Worksheets("WBS")
    Set tbl = ws.ListObjects("tbl_WBS")
    If Trim$(languageCode) <> "" Then
        WBS_SetLanguage languageCode
    Else
        WBS_SetLanguage Settings_GetOwnerLanguage("WBS")
    End If
    WBS_OnboardingHarness_CheckLocalized = CStr( _
        WBS_Onboarding_LocalizedContentIsCurrent(ws, tbl))
End Function

Public Function WBS_OnboardingHarness_ApplyLocalized() As String
    Dim ws As Worksheet
    Dim tbl As ListObject

    Set ws = ThisWorkbook.Worksheets("WBS")
    Set tbl = ws.ListObjects("tbl_WBS")
    WBS_SetLanguage Settings_GetOwnerLanguage("WBS")
    WBS_Onboarding_ApplyLocalizedContent ws, tbl
    WBS_OnboardingHarness_ApplyLocalized = "PASS"
End Function

'------------------------------------------------------------------------------
' FR: Verifie en lecture seule que la ligne d'onboarding utilise la structure canonique.
' EN: Read-only check that the onboarding row uses the canonical structure.
'------------------------------------------------------------------------------
Private Function WBS_Onboarding_StructureIsCurrent( _
    ByVal ws As Worksheet, _
    ByVal tbl As ListObject, _
    ByVal statuses As Object) As Boolean

    Dim key As Variant
    Dim cell As Range


    On Error GoTo NotCurrent

    If ws Is Nothing Or tbl Is Nothing Or statuses Is Nothing Then Exit Function
    If tbl.HeaderRowRange.Row <= 1 Then Exit Function
    If Not tbl.DataBodyRange Is Nothing Then
        If tbl.DataBodyRange.Row <> tbl.HeaderRowRange.Row + 1 Then Exit Function
    End If
    If Abs(CDbl(ws.Rows(WBS_OnboardingHelpRow(tbl)).RowHeight) - _
        WBS_ONBOARDING_ROW_HEIGHT) > 0.05 Then Exit Function

    For Each key In statuses.Keys
        Set cell = ws.Cells(WBS_OnboardingHelpRow(tbl), _
            SchemaListColumn(tbl, VTS_TABLE_WBS, CStr(key)).Range.Column)


        If CStr(cell.Font.Name) <> WBS_ONBOARDING_FONT_NAME Then Exit Function
        If Abs(CDbl(cell.Font.Size) - WBS_ONBOARDING_FONT_SIZE) > 0.01 Then Exit Function
        If Not CBool(cell.Font.Bold) Then Exit Function
        If CBool(cell.Font.Italic) Then Exit Function
        If cell.Interior.Pattern <> xlSolid Then Exit Function
        If cell.Interior.Color <> RGB(255, 255, 255) Then Exit Function
        If cell.HorizontalAlignment <> xlCenter Then Exit Function
        If cell.VerticalAlignment <> xlCenter Then Exit Function
        If Not CBool(cell.WrapText) Then Exit Function
        If CBool(cell.ShrinkToFit) Then Exit Function
        If cell.Borders(xlEdgeBottom).LineStyle <> xlContinuous Then Exit Function
        If cell.Borders(xlEdgeBottom).Color <> RGB(0, 0, 0) Then Exit Function
        If cell.Borders(xlEdgeBottom).Weight <> xlThin Then Exit Function

    Next key

    WBS_Onboarding_StructureIsCurrent = True
    Exit Function

NotCurrent:
    WBS_Onboarding_StructureIsCurrent = False

End Function

'------------------------------------------------------------------------------
' FR: Repare la ligne d'onboarding lorsqu'un ecart structurel est demontre.
' EN: Repairs the onboarding row when a structural mismatch is proven.
'------------------------------------------------------------------------------
Private Sub WBS_Onboarding_ApplyStructure( _
    ByVal ws As Worksheet, _
    ByVal tbl As ListObject, _
    ByVal statuses As Object)

    Dim helpRange As Range
    Dim key As Variant
    Dim cell As Range

    Set helpRange = ws.Range( _
        ws.Cells(WBS_OnboardingHelpRow(tbl), tbl.Range.Column), _
        ws.Cells(WBS_OnboardingHelpRow(tbl), _
            tbl.Range.Column + tbl.ListColumns.Count - 1))

    helpRange.ClearContents
    With helpRange
        .Font.Name = WBS_ONBOARDING_FONT_NAME
        .Font.Size = WBS_ONBOARDING_FONT_SIZE
        .Font.Bold = True
        .Font.Italic = False
        .Font.Color = RGB(0, 0, 0)
        .Interior.Pattern = xlSolid
        .Interior.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .ShrinkToFit = False
        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Color = RGB(0, 0, 0)
            .Weight = xlThin
        End With
    End With

    ws.Rows(WBS_OnboardingHelpRow(tbl)).RowHeight = WBS_ONBOARDING_ROW_HEIGHT

    For Each key In statuses.Keys
        Set cell = ws.Cells(WBS_OnboardingHelpRow(tbl), _
            SchemaListColumn(tbl, VTS_TABLE_WBS, CStr(key)).Range.Column)
        cell.Value = CStr(statuses(key))
        WBS_Onboarding_FormatStatusCell cell
    Next key

End Sub

'------------------------------------------------------------------------------
' FR: Reconstruit les statuts, le Quick Start et les commentaires dans la langue active.
' EN: Rebuilds statuses, the Quick Start and comments in the active language.
'------------------------------------------------------------------------------
Private Sub WBS_Onboarding_ApplyLocalizedContent( _
    ByVal ws As Worksheet, _
    Optional ByVal tbl As ListObject = Nothing)

    Dim comments As Object
    Dim statuses As Object
    Dim key As Variant
    Dim statusCell As Range
    Dim listColumn As ListColumn
    Dim helpCell As Range
    Dim headerCell As Range
    Dim localizedPair As Variant
    Dim expectedText As String
    Dim actualText As String
    Dim hasThreadedComment As Boolean
    Dim threadedCommentsSupported As Boolean
    Dim languageIndex As Long
    Dim columnKeys As Variant
    Dim oldEvents As Boolean
    Dim oldScreenUpdating As Boolean
    Dim errorNumber As Long
    Dim errorSource As String
    Dim errorDescription As String

    If ws Is Nothing Then Exit Sub

    oldEvents = Application.EnableEvents
    oldScreenUpdating = Application.ScreenUpdating
    On Error GoTo ErrHandler

    If tbl Is Nothing Then Set tbl = ws.ListObjects("tbl_WBS")
    Set comments = WBS_Onboarding_BuildHelpCommentMap()
    Set statuses = WBS_Onboarding_BuildStatusMap()
    threadedCommentsSupported = ExcelCompatibility_SupportsThreadedComments()
    languageIndex = IIf(WBS_CurrentLanguage() = "FR", 0, 1)
    columnKeys = SchemaColumnKeys(VTS_TABLE_WBS)

    Application.EnableEvents = False
    Application.ScreenUpdating = False
    gWBSOnboardingLocalizedRebuildCount = _
        gWBSOnboardingLocalizedRebuildCount + 1

    For Each key In statuses.Keys
        Set statusCell = ws.Cells(WBS_OnboardingHelpRow(tbl), _
            SchemaListColumn(tbl, VTS_TABLE_WBS, CStr(key)).Range.Column)
        If CStr(statusCell.Value2) <> CStr(statuses(key)) Then
            statusCell.Value = CStr(statuses(key))
        End If
        WBS_Onboarding_FormatStatusCell statusCell
    Next key

    For Each key In columnKeys
        Set listColumn = SchemaListColumn(tbl, VTS_TABLE_WBS, CStr(key))
        Set helpCell = ws.Cells(WBS_OnboardingHelpRow(tbl), listColumn.Range.Column)
        Set headerCell = listColumn.Range.Cells(1, 1)

        If WBS_Onboarding_CellHasAnyComment(headerCell) Then
            WBS_Onboarding_ClearCellComments headerCell
        End If

        If comments.Exists(CStr(key)) Then
            localizedPair = comments(CStr(key))
            expectedText = CStr(localizedPair(languageIndex))
            If threadedCommentsSupported Then
                actualText = WBS_Onboarding_ThreadedCommentText( _
                    helpCell, hasThreadedComment)
                If WBS_Onboarding_CellHasLegacyComment(helpCell) Or _
                   Not hasThreadedComment Or _
                   WBS_Onboarding_NormalizeComparisonText(actualText) <> _
                        WBS_Onboarding_NormalizeComparisonText(expectedText) Then
                    WBS_Onboarding_ClearCellComments helpCell
                    WBS_Onboarding_AddThreadedComment helpCell, expectedText
                End If
            Else
                actualText = WBS_Onboarding_LegacyCommentText( _
                    helpCell, hasThreadedComment)
                If Not hasThreadedComment Or _
                   WBS_Onboarding_NormalizeComparisonText(actualText) <> _
                        WBS_Onboarding_NormalizeComparisonText(expectedText) Then
                    WBS_Onboarding_ClearCellComments helpCell
                    WBS_Onboarding_AddLegacyComment helpCell, expectedText
                End If
            End If
        ElseIf WBS_Onboarding_CellHasAnyComment(helpCell) Then
            WBS_Onboarding_ClearCellComments helpCell
        End If
    Next key

    WBS_Onboarding_WriteQuickStart ws

    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEvents
    Exit Sub

ErrHandler:
    errorNumber = Err.Number
    errorSource = Err.Source
    errorDescription = Err.Description

    On Error Resume Next
    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEvents
    On Error GoTo 0

    Err.Raise errorNumber, errorSource, errorDescription

End Sub

'------------------------------------------------------------------------------
' FR: Verifie en lecture seule que les contenus localises correspondent au catalogue actif.
' EN: Read-only check that localized content matches the active catalog.
'------------------------------------------------------------------------------
Private Function WBS_Onboarding_LocalizedContentIsCurrent( _
    ByVal ws As Worksheet, _
    Optional ByVal tbl As ListObject = Nothing) As Boolean

    Dim comments As Object
    Dim statuses As Object
    Dim key As Variant
    Dim statusCell As Range
    Dim requiredFont As Font
    Dim suffixFont As Font
    Dim requiredLabel As String
    Dim listColumn As ListColumn
    Dim helpCell As Range
    Dim headerCell As Range
    Dim noteArea As Range
    Dim localizedPair As Variant
    Dim expectedText As String
    Dim actualText As String
    Dim hasThreadedComment As Boolean
    Dim threadedCommentsSupported As Boolean
    Dim languageIndex As Long
    Dim columnKeys As Variant

    gWBSOnboardingLocalizedCheckCount = _
        gWBSOnboardingLocalizedCheckCount + 1

    On Error GoTo NotCurrent
    If ws Is Nothing Then Exit Function
    If tbl Is Nothing Then Set tbl = ws.ListObjects("tbl_WBS")

    Set comments = WBS_Onboarding_BuildHelpCommentMap()
    Set statuses = WBS_Onboarding_BuildStatusMap()
    threadedCommentsSupported = ExcelCompatibility_SupportsThreadedComments()
    requiredLabel = WBS_Onboarding_RequiredLabel()
    If comments.Count <> 37 Then Exit Function
    If Not comments.Exists(VTS_COL_PROJECT) Then Exit Function
    If Not comments.Exists(VTS_COL_LONGEST_PATH) Then Exit Function
    If Not comments.Exists(VTS_COL_LONGEST_PATH_REX) Then Exit Function
    If Not comments.Exists(VTS_COL_DEADLINE_FLOAT) Then Exit Function

    languageIndex = IIf(WBS_CurrentLanguage() = "FR", 0, 1)
    columnKeys = SchemaColumnKeys(VTS_TABLE_WBS)
    Set noteArea = ws.Range("O1:R2")

    If Not CBool(noteArea.MergeCells) Then Exit Function
    If ws.Range("O1").MergeArea.Address(False, False) <> "O1:R2" Then Exit Function
    If WBS_Onboarding_NormalizeComparisonText(CStr(ws.Range("O1").Value2)) <> _
        WBS_Onboarding_NormalizeComparisonText(WBS_Onboarding_QuickStartText()) Then Exit Function
    If noteArea.HorizontalAlignment <> xlLeft Then Exit Function
    If noteArea.VerticalAlignment <> xlTop Then Exit Function
    If Not CBool(noteArea.WrapText) Then Exit Function

    For Each key In statuses.Keys
        Set statusCell = ws.Cells(WBS_OnboardingHelpRow(tbl), _
            SchemaListColumn(tbl, VTS_TABLE_WBS, CStr(key)).Range.Column)
        actualText = CStr(statusCell.Value2)
        If actualText <> CStr(statuses(key)) Then Exit Function

        If Left$(actualText, Len(requiredLabel)) = requiredLabel Then
            Set requiredFont = statusCell.Characters(1, Len(requiredLabel)).Font
            If requiredFont.Color <> RGB(192, 0, 0) Then Exit Function
            If Not CBool(requiredFont.Bold) Then Exit Function

            If Len(actualText) > Len(requiredLabel) Then
                Set suffixFont = statusCell.Characters( _
                    Len(requiredLabel) + 1, _
                    Len(actualText) - Len(requiredLabel)).Font
                If suffixFont.Color <> RGB(0, 0, 0) Then Exit Function
            End If
        ElseIf statusCell.Font.Color <> RGB(0, 0, 0) Then
            Exit Function
        End If
    Next key

    For Each key In columnKeys
        Set listColumn = SchemaListColumn(tbl, VTS_TABLE_WBS, CStr(key))
        Set helpCell = ws.Cells(WBS_OnboardingHelpRow(tbl), listColumn.Range.Column)
        Set headerCell = listColumn.Range.Cells(1, 1)

        If WBS_Onboarding_CellHasAnyComment(headerCell) Then Exit Function

        If threadedCommentsSupported Then
            If WBS_Onboarding_CellHasLegacyComment(helpCell) Then Exit Function
            actualText = WBS_Onboarding_ThreadedCommentText( _
                helpCell, hasThreadedComment)
        Else
            actualText = WBS_Onboarding_LegacyCommentText( _
                helpCell, hasThreadedComment)
        End If

        If comments.Exists(CStr(key)) Then
            If Not hasThreadedComment Then Exit Function
            localizedPair = comments(CStr(key))
            expectedText = CStr(localizedPair(languageIndex))
            If WBS_Onboarding_NormalizeComparisonText(actualText) <> _
                WBS_Onboarding_NormalizeComparisonText(expectedText) Then Exit Function
        ElseIf hasThreadedComment Then
            Exit Function
        End If
    Next key

    WBS_Onboarding_LocalizedContentIsCurrent = True
    Exit Function

NotCurrent:
    WBS_Onboarding_LocalizedContentIsCurrent = False

End Function

'------------------------------------------------------------------------------
' FR: Retourne le Quick Start canonique dans la langue runtime WBS.
' EN: Returns the canonical Quick Start in the WBS runtime language.
'------------------------------------------------------------------------------
Private Function WBS_Onboarding_QuickStartText() As String

    WBS_Onboarding_QuickStartText = WBS_L( _
        WBS_QUICK_START_FR_LINE_1 & vbLf & WBS_QUICK_START_FR_LINE_2, _
        WBS_QUICK_START_EN_LINE_1 & vbLf & WBS_QUICK_START_EN_LINE_2)

End Function

'------------------------------------------------------------------------------
' FR: Normalise uniquement les fins de ligne avant une comparaison de contenu.
' EN: Normalizes line endings only before comparing content.
'------------------------------------------------------------------------------
Private Function WBS_Onboarding_NormalizeComparisonText( _
    ByVal value As String) As String

    value = Replace(value, vbCrLf, vbLf)
    value = Replace(value, vbCr, vbLf)
    WBS_Onboarding_NormalizeComparisonText = value

End Function

'------------------------------------------------------------------------------
' FR: Indique si une cellule possede un commentaire classique.
' EN: Reports whether a cell owns a legacy comment.
'------------------------------------------------------------------------------
Private Function WBS_Onboarding_CellHasLegacyComment( _
    ByVal cell As Range) As Boolean

    Dim legacyComment As Object

    On Error Resume Next
    Set legacyComment = cell.Comment
    On Error GoTo 0
    WBS_Onboarding_CellHasLegacyComment = Not legacyComment Is Nothing

End Function

'------------------------------------------------------------------------------
' FR: Lit une Note Excel classique sans utiliser l'API threaded.
' EN: Reads a classic Excel Note without using the threaded API.
'------------------------------------------------------------------------------
Private Function WBS_Onboarding_LegacyCommentText( _
    ByVal cell As Range, _
    ByRef commentExists As Boolean) As String

    Dim legacyComment As Object

    commentExists = False
    On Error Resume Next
    Set legacyComment = cell.Comment
    On Error GoTo 0
    If legacyComment Is Nothing Then Exit Function

    commentExists = True
    WBS_Onboarding_LegacyCommentText = CStr( _
        CallByName(legacyComment, "Text", VbMethod))

End Function

'------------------------------------------------------------------------------
' FR: Lit un commentaire threaded sans modifier la cellule.
' EN: Reads a threaded comment without mutating the cell.
'------------------------------------------------------------------------------
Private Function WBS_Onboarding_ThreadedCommentText( _
    ByVal cell As Range, _
    ByRef commentExists As Boolean) As String

    Dim threadedComment As Object

    commentExists = False
    If Not ExcelCompatibility_SupportsThreadedComments() Then Exit Function

    gWBSOnboardingThreadedReadCount = gWBSOnboardingThreadedReadCount + 1
    On Error Resume Next
    Set threadedComment = cell.CommentThreaded
    On Error GoTo 0

    If threadedComment Is Nothing Then Exit Function
    commentExists = True
    WBS_Onboarding_ThreadedCommentText = CStr( _
        CallByName(threadedComment, "Text", VbMethod))

End Function

'------------------------------------------------------------------------------
' FR: Indique si une cellule possede un commentaire classique ou threaded.
' EN: Reports whether a cell owns a legacy or threaded comment.
'------------------------------------------------------------------------------
Private Function WBS_Onboarding_CellHasAnyComment( _
    ByVal cell As Range) As Boolean

    Dim hasThreadedComment As Boolean
    Dim ignoredText As String

    ignoredText = WBS_Onboarding_ThreadedCommentText(cell, hasThreadedComment)
    WBS_Onboarding_CellHasAnyComment = _
        hasThreadedComment Or WBS_Onboarding_CellHasLegacyComment(cell)

End Function

'------------------------------------------------------------------------------
' FR: Supprime les commentaires d'une cellule uniquement lorsqu'une reparation est requise.
' EN: Clears cell comments only when a repair is required.
'------------------------------------------------------------------------------
Private Sub WBS_Onboarding_ClearCellComments(ByVal cell As Range)

    Dim legacyComment As Object
    Dim threadedComment As Object

    On Error Resume Next
    Set legacyComment = cell.Comment
    If Not legacyComment Is Nothing Then legacyComment.Delete
    If ExcelCompatibility_SupportsThreadedComments() Then
        Set threadedComment = cell.CommentThreaded
        If Not threadedComment Is Nothing Then
            gWBSOnboardingThreadedClearCount = gWBSOnboardingThreadedClearCount + 1
            CallByName threadedComment, "Delete", VbMethod
        End If
    End If
    On Error GoTo 0

End Sub

'------------------------------------------------------------------------------
' FR: Ajoute un commentaire threaded uniquement lorsque l'API Excel est disponible.
' EN: Adds a threaded comment only when the Excel API is available.
'------------------------------------------------------------------------------
Private Sub WBS_Onboarding_AddThreadedComment( _
    ByVal cell As Range, _
    ByVal commentText As String)

    If Not ExcelCompatibility_SupportsThreadedComments() Then Exit Sub

    gWBSOnboardingThreadedAddCount = gWBSOnboardingThreadedAddCount + 1
    CallByName cell, "AddCommentThreaded", VbMethod, commentText

End Sub

'------------------------------------------------------------------------------
' FR: Ajoute une Note Excel classique sur les versions sans commentaires threaded.
' EN: Adds a classic Excel Note on versions without threaded comments.
'------------------------------------------------------------------------------
Private Sub WBS_Onboarding_AddLegacyComment( _
    ByVal cell As Range, _
    ByVal commentText As String)

    cell.AddComment commentText

End Sub

'------------------------------------------------------------------------------
' FR: Construit le catalogue canonique des aides de colonnes WBS en FR et EN.
' EN: Builds the canonical FR and EN catalog of WBS column help text.
'------------------------------------------------------------------------------
Private Function WBS_Onboarding_BuildHelpCommentMap() As Object

    Dim comments As Object

    Set comments = CreateObject("Scripting.Dictionary")
    comments.CompareMode = vbTextCompare

    WBS_Onboarding_AddHelpComment comments, VTS_COL_ID, _
        "Rôle : identifiant unique de la ligne.{NL}Saisie attendue : entier unique, sans doublon.{NL}Exemple : 1 / 17 / 29{NL}{NL}Utilité :{NL}- Sert de clé technique entre WBS, CALC, GANTT et S-Curve.{NL}- Ne doit jamais être dupliqué.{NL}- Peut rester simple même si le WBS change.", _
        "Purpose: unique row identifier.{NL}Expected input: unique integer, no duplicates.{NL}Example: 1 / 17 / 29{NL}{NL}Use:{NL}- Technical key used across WBS, CALC, GANTT and S-Curve.{NL}- Must never be duplicated.{NL}- Can stay stable even if the WBS code changes."
    WBS_Onboarding_AddHelpComment comments, VTS_COL_WBS, _
        "Rôle : code hiérarchique de la tâche dans la structure du projet.{NL}Saisie attendue : format numérique hiérarchique avec points.{NL}Exemple : 1.0 / 1.3 / 1.3.2 / 1.3.2.1{NL}{NL}Règles :{NL}- Utiliser uniquement des chiffres et des points.{NL}- Pas d’espace, pas de lettres.{NL}- Une tâche parent a des enfants dont le WBS commence par son propre code.", _
        "Purpose: hierarchical code of the task in the project structure.{NL}Expected input: numeric hierarchical format using dots.{NL}Example: 1.0 / 1.3 / 1.3.2 / 1.3.2.1{NL}{NL}Rules:{NL}- Use numbers and dots only.{NL}- No spaces, no letters.{NL}- A parent task has child tasks whose WBS starts with its own code."
    WBS_Onboarding_AddHelpComment comments, VTS_COL_TASK_NAME, _
        "Rôle : nom court de la tâche.{NL}Saisie attendue : intitulé clair et lisible.{NL}Exemple : Kick-off / RFQ / Assembly / FAT{NL}{NL}Conseil :{NL}- Rester court.{NL}- Utiliser un nom orienté action ou livrable.", _
        "Purpose: short task name.{NL}Expected input: clear readable label.{NL}Example: Kick-off / RFQ / Assembly / FAT{NL}{NL}Tip:{NL}- Keep it short.{NL}- Prefer action-oriented or deliverable-oriented wording."
    WBS_Onboarding_AddHelpComment comments, VTS_COL_TASK_DESCRIPTION, _
        "Rôle : description détaillée de la tâche.{NL}Saisie attendue : phrase courte ou précision utile.{NL}Exemple : Project official start meeting / Vendor docs review{NL}{NL}Utilité :{NL}- Aide à comprendre la tâche sans lire tout le planning.", _
        "Purpose: detailed task description.{NL}Expected input: short sentence or useful clarification.{NL}Example: Project official start meeting / Vendor docs review{NL}{NL}Use:{NL}- Helps understand the task without reading the full schedule."
    WBS_Onboarding_AddHelpComment comments, VTS_COL_DISCIPLINE, _
        "Rôle : discipline technique principale concernée.{NL}Saisie attendue : nom de discipline cohérent sur tout le fichier.{NL}Exemple : Process / Mechanical / Structure / QAQC / Logistics{NL}{NL}Conseil :{NL}- Garder un vocabulaire homogène dans tout le planning.", _
        "Purpose: main technical discipline involved.{NL}Expected input: discipline name used consistently across the file.{NL}Example: Process / Mechanical / Structure / QAQC / Logistics{NL}{NL}Tip:{NL}- Keep the wording consistent throughout the schedule."
    WBS_Onboarding_AddHelpComment comments, VTS_COL_SUPPLIER, _
        "Rôle : acteur principal responsable ou concerné par la tâche.{NL}Saisie attendue : nom de société, fournisseur ou entité interne.{NL}Exemple : Internal / Vendor A / Client / Forwarder", _
        "Purpose: main party responsible for or involved in the task.{NL}Expected input: company name, supplier or internal entity.{NL}Example: Internal / Vendor A / Client / Forwarder"
    WBS_Onboarding_AddHelpComment comments, VTS_COL_PROJECT, _
        "Rôle : projet ou regroupement métier facultatif associé à la tâche.{NL}Saisie attendue : nom du projet ou regroupement utile.{NL}Exemple : Projet A / Fabrication / Zone 2{NL}{NL}Utilité :{NL}- Facilite le regroupement métier des tâches.{NL}- Peut rester vide.{NL}- N'intervient pas dans le calcul du planning.", _
        "Purpose: optional project or business grouping associated with the task.{NL}Expected input: project name or useful grouping.{NL}Example: Project A / Fabrication / Area 2{NL}{NL}Use:{NL}- Helps group tasks for business reporting.{NL}- May be left blank.{NL}- Does not affect schedule calculation."
    WBS_Onboarding_AddHelpComment comments, VTS_COL_TASK_TYPE, _
        "Rôle : définit le comportement de la tâche dans le moteur de planning.{NL}{NL}Valeurs autorisées :{NL}Task{NL}Milestone{NL}Level of Effort{NL}{NL}Définition :{NL}{NL}Task{NL}tâche standard avec durée{NL}peut avoir des prédécesseurs{NL}pilotée par Actual / Forecast / Baseline / Dependencies{NL}impacte le réseau normalement{NL}{NL}Milestone{NL}tâche sans durée (événement ponctuel){NL}durée nulle ou minimale selon convention{NL}peut avoir des prédécesseurs{NL}représente un jalon (début / fin / validation){NL}{NL}Level of Effort{NL}tâche dépendante d’une plage d’activités{NL}ne pilote pas le planning{NL}est pilotée par ses dépendances{NL}généralement définie par SS (début) et FF (fin){NL}durée déduite du réseau{NL}{NL}Règles :{NL}une seule valeur par tâche{NL}respecter strictement les valeurs autorisées{NL}une LOE ne doit pas être utilisée comme driver d’autres tâches{NL}une milestone ne doit pas porter de durée métier", _
        "Purpose: defines the task behavior within the planning engine.{NL}{NL}Allowed values:{NL}Task{NL}Milestone{NL}Level of Effort{NL}{NL}Definition:{NL}{NL}Task{NL}standard task with duration{NL}can have predecessors{NL}driven by Actual / Forecast / Baseline / Dependencies{NL}fully participates in the network{NL}{NL}Milestone{NL}zero-duration task (event){NL}duration is zero or minimal depending on convention{NL}can have predecessors{NL}represents a key event (start / finish / validation){NL}{NL}Level of Effort{NL}task spanning a range of activities{NL}does not drive the schedule{NL}is driven by its dependencies{NL}typically defined using SS (start) and FF (finish) links{NL}duration is derived from the network{NL}{NL}Rules:{NL}one value per task{NL}must match allowed values exactly{NL}a LOE must not be used as a driver for other tasks{NL}a milestone must not carry business duration"
    WBS_Onboarding_AddHelpComment comments, VTS_COL_S, _
        "Rôle : définit si la tâche doit apparaître dans la vue Summary du Gantt.{NL}Valeurs autorisées :{NL}Y{NL}N{NL}Définition :{NL}Y{NL}la tâche est affichée dans la vue Summary{NL}peut être utilisé pour afficher une tâche standard importante{NL}permet de forcer l’affichage d’une ligne même si ce n’est pas un parent ou une milestone{NL}N{NL}la tâche est masquée dans la vue Summary{NL}permet de masquer une milestone ou une ligne non pertinente{NL}n’impacte pas le calcul planning{NL}Règles :{NL}si vide, la valeur est remplie automatiquement{NL}parents / summaries : Y par défaut{NL}milestones : Y par défaut{NL}tasks standard : N par défaut{NL}Level of Effort : N par défaut{NL}une valeur déjà renseignée n’est jamais écrasée{NL}Y ou N uniquement", _
        "Purpose: defines whether the task should appear in the Gantt Summary view.{NL}Allowed values:{NL}Y{NL}N{NL}Definition:{NL}Y{NL}task is displayed in the Summary view{NL}can be used to show an important standard task{NL}forces a row to appear even if it is not a parent or milestone{NL}N{NL}task is hidden from the Summary view{NL}can be used to hide a milestone or non-relevant row{NL}does not impact schedule calculation{NL}Rules:{NL}if blank, the value is filled automatically{NL}parents / summaries: Y by default{NL}milestones: Y by default{NL}standard tasks: N by default{NL}Level of Effort: N by default{NL}an existing value is never overwritten{NL}Y or N only"
    WBS_Onboarding_AddHelpComment comments, VTS_COL_COMMENTS, _
        "Rôle : zone libre pour note de contexte.{NL}Saisie attendue : commentaire court, risque, hypothèse, précision.{NL}Exemple : Waiting vendor confirmation / Milestone imposed by client", _
        "Purpose: free text field for context notes.{NL}Expected input: short comment, risk, assumption or clarification.{NL}Example: Waiting vendor confirmation / Milestone imposed by client"
    WBS_Onboarding_AddHelpComment comments, VTS_COL_PREDECESSORS_WBS, _
        "Rôle : antécédents de la tâche, saisis au format WBS avec type de lien et lag éventuel.{NL}Saisie attendue :{NL}un ou plusieurs prédécesseurs séparés par un point-virgule{NL}type par défaut = FS si rien n’est précisé{NL}Exemples :{NL}1.2.3{NL}1.2.3+4{NL}1.2.3-2{NL}1.2.3FS+4{NL}1.2.3SS-2{NL}1.2.3FF{NL}1.2.3;1.4.1SS+2;2.3FF-1{NL}Règles :{NL}pas d’espace{NL}utiliser le WBS, pas l’ID{NL}types autorisés : FS, SS, FF{NL}lag autorisé en positif ou négatif{NL}le moteur convertit ensuite cette donnée en IDs techniques + table de liens logiques", _
        "Purpose: task predecessors entered using WBS codes, with optional link type and lag.{NL}Expected input:{NL}one or more predecessors separated by semicolons{NL}default link type = FS when omitted{NL}Examples:{NL}1.2.3{NL}1.2.3+4{NL}1.2.3-2{NL}1.2.3FS+4{NL}1.2.3SS-2{NL}1.2.3FF{NL}1.2.3;1.4.1SS+2;2.3FF-1{NL}Rules:{NL}no spaces{NL}use WBS, not ID{NL}allowed link types: FS, SS, FF{NL}positive or negative lag allowed{NL}the engine then converts this input into technical IDs + logical link table"
    WBS_Onboarding_AddHelpComment comments, VTS_COL_WEIGHT_PERCENT, _
        "Rôle : poids de la tâche pour les analyses de charge ou de progression pondérée.{NL}Saisie attendue : valeur de poids selon la logique projet.{NL}Exemple : 20000 € / 15 / 4.5{NL}{NL}Utilité :{NL}- Peut représenter un coût, une charge, un volume ou tout autre poids relatif.{NL}- La S-Curve travaille sur les tâches feuilles uniquement et normalise ensuite les poids.", _
        "Purpose: task weight used for workload or weighted progress analysis.{NL}Expected input: weight value according to the project logic.{NL}Example: 20000 € / 15 / 4.5{NL}{NL}Use:{NL}- Can represent cost, effort, quantity or any relative weighting.{NL}- The S-Curve works on leaf tasks only and then normalizes weights."
    WBS_Onboarding_AddHelpComment comments, VTS_COL_PROGRESS_PERCENT, _
        "Rôle : avancement manuel de la tâche.{NL}Saisie attendue : pourcentage entre 0% et 100%.{NL}Exemple : 0% / 8% / 70% / 100%{NL}{NL}Règles :{NL}- À renseigner sur les tâches feuilles.{NL}- En l’absence de valeur, l’affichage Gantt peut considérer 0%.", _
        "Purpose: manual task progress.{NL}Expected input: percentage between 0% and 100%.{NL}Example: 0% / 8% / 70% / 100%{NL}{NL}Rules:{NL}- Meant for leaf tasks.{NL}- When empty, Gantt display may treat it as 0%."
    WBS_Onboarding_AddHelpComment comments, VTS_COL_BASELINE_START, _
        "Rôle : date de début de référence.{NL}Saisie attendue : date baseline prévue au plan initial.{NL}Exemple : 05/02/2026{NL}{NL}Utilité :{NL}- Sert de base de comparaison pour les écarts.{NL}- Utilisée par le moteur et par les analyses REX.", _
        "Purpose: reference start date.{NL}Expected input: baseline start date from the initial plan.{NL}Example: 05/02/2026{NL}{NL}Use:{NL}- Used as comparison basis for variances.{NL}- Used by the engine and by REX analyses."
    WBS_Onboarding_AddHelpComment comments, VTS_COL_BASELINE_DURATION, _
        "Rôle : durée baseline en jours calendaires inclusifs.{NL}Saisie attendue : entier positif.{NL}Exemple : 1 / 5 / 12{NL}{NL}Règle importante :{NL}- Une durée de 1 jour signifie début = fin le même jour.", _
        "Purpose: baseline duration in inclusive calendar days.{NL}Expected input: positive integer.{NL}Example: 1 / 5 / 12{NL}{NL}Important rule:{NL}- A duration of 1 day means start = finish on the same day."
    WBS_Onboarding_AddHelpComment comments, VTS_COL_BASELINE_FINISH, _
        "Rôle : date de fin baseline.{NL}Calcul / logique :{NL}- Colonne calculée automatiquement à partir de Baseline Start et Baseline Duration.{NL}- Logique inclusive : Finish = Start + Duration - 1{NL}{NL}Utilité :{NL}- Sert aux écarts de fin et aux comparaisons planning.", _
        "Purpose: baseline finish date.{NL}Calculation / logic:{NL}- Automatically calculated from Baseline Start and Baseline Duration.{NL}- Inclusive logic: Finish = Start + Duration - 1{NL}{NL}Use:{NL}- Used for finish variance and schedule comparisons."
    WBS_Onboarding_AddHelpComment comments, VTS_COL_ACTUAL_START, _
        "Rôle : date de début réellement constatée.{NL}Saisie attendue : date réelle si la tâche a commencé.{NL}Exemple : 21/03/2026{NL}{NL}Utilité :{NL}- Prioritaire sur Forecast et Baseline pour le calcul moteur.", _
        "Purpose: actual observed start date.{NL}Expected input: real start date if the task has started.{NL}Example: 21/03/2026{NL}{NL}Use:{NL}- Has priority over Forecast and Baseline in the engine logic."
    WBS_Onboarding_AddHelpComment comments, VTS_COL_ACTUAL_FINISH, _
        "Rôle : date de fin réellement constatée.{NL}Saisie attendue : date réelle si la tâche est terminée.{NL}Exemple : 23/03/2026{NL}{NL}Utilité :{NL}- Prioritaire pour le calcul de la fin si présente.", _
        "Purpose: actual observed finish date.{NL}Expected input: real finish date if the task is completed.{NL}Example: 23/03/2026{NL}{NL}Use:{NL}- Has priority for finish calculation when present."
    WBS_Onboarding_AddHelpComment comments, VTS_COL_ACTUAL_DURATION, _
        "Rôle : durée réelle observée.{NL}Calcul / logique :{NL}- Colonne calculée automatiquement à partir de Actual Start et Actual Finish.{NL}- Logique inclusive : Duration = Finish - Start + 1{NL}{NL}Utilité :{NL}- Donne la durée réelle constatée sans saisie manuelle.", _
        "Purpose: actual observed duration.{NL}Calculation / logic:{NL}- Automatically calculated from Actual Start and Actual Finish.{NL}- Inclusive logic: Duration = Finish - Start + 1{NL}{NL}Use:{NL}- Provides the real duration without manual entry."
    WBS_Onboarding_AddHelpComment comments, VTS_COL_FORECAST_START, _
        "Rôle : date de début prévisionnelle mise à jour.{NL}Saisie attendue : date forecast si la tâche n’est pas entièrement portée par l’Actual.{NL}Exemple : 04/04/2026{NL}{NL}Utilité :{NL}- Permet de simuler ou piloter une dérive planning.{NL}- Si incohérente avec les dépendances, le moteur bloque.", _
        "Purpose: updated forecast start date.{NL}Expected input: forecast date when the task is not fully driven by Actual data.{NL}Example: 04/04/2026{NL}{NL}Use:{NL}- Allows schedule drift management and simulation.{NL}- If inconsistent with dependencies, the engine blocks."
    WBS_Onboarding_AddHelpComment comments, VTS_COL_FORECAST_FINISH, _
        "Rôle : date de fin prévisionnelle mise à jour.{NL}Saisie attendue : date forecast de fin.{NL}Exemple : 30/04/2026{NL}{NL}Utilité :{NL}- Permet d’imposer une fin forecast.{NL}- Si seule la date de début est donnée, le moteur conserve la durée de référence.", _
        "Purpose: updated forecast finish date.{NL}Expected input: forecast finish date.{NL}Example: 30/04/2026{NL}{NL}Use:{NL}- Allows forcing a forecast finish.{NL}- If only the start is given, the engine keeps the reference duration."
    WBS_Onboarding_AddHelpComment comments, VTS_COL_CALCULATED_START, _
        "Rôle : date de début calculée par le moteur.{NL}Calcul / logique :{NL}- Priorité générale : Actual > Forecast > Baseline > Dépendances seules{NL}- Le moteur tient compte des prédécesseurs et du lag.{NL}{NL}Utilité :{NL}- Référence consolidée utilisée pour le Gantt et les analyses.", _
        "Purpose: engine-calculated start date.{NL}Calculation / logic:{NL}- General priority: Actual > Forecast > Baseline > Dependencies only{NL}- The engine also applies predecessors and lag.{NL}{NL}Use:{NL}- Consolidated reference used by Gantt and analyses."
    WBS_Onboarding_AddHelpComment comments, VTS_COL_CALCULATED_FINISH, _
        "Rôle : date de fin calculée par le moteur.{NL}Calcul / logique :{NL}- Basée sur Actual Finish si présent, sinon Forecast Finish si présent, sinon durée de référence.{NL}- Toujours cohérente avec Calculated Start si le calcul réussit.{NL}{NL}Utilité :{NL}- Référence consolidée utilisée pour le Gantt et les analyses.", _
        "Purpose: engine-calculated finish date.{NL}Calculation / logic:{NL}- Based on Actual Finish if present, otherwise Forecast Finish if present, otherwise reference duration.{NL}- Always aligned with Calculated Start if the calculation succeeds.{NL}{NL}Use:{NL}- Consolidated reference used by Gantt and analyses."
    WBS_Onboarding_AddHelpComment comments, VTS_COL_CALCULATED_DURATION, _
        "Rôle : durée calculée consolidée.{NL}Calcul / logique :{NL}- Colonne calculée automatiquement à partir de Calculated Start et Calculated Finish.{NL}- Logique inclusive : Duration = Finish - Start + 1{NL}{NL}Utilité :{NL}- Affiche la durée réellement retenue après calcul.", _
        "Purpose: consolidated calculated duration.{NL}Calculation / logic:{NL}- Automatically calculated from Calculated Start and Calculated Finish.{NL}- Inclusive logic: Duration = Finish - Start + 1{NL}{NL}Use:{NL}- Shows the duration finally retained after calculation."
    WBS_Onboarding_AddHelpComment comments, VTS_COL_START_VARIANCE, _
        "Rôle : écart entre le début calculé et le début baseline.{NL}Calcul / logique :{NL}- Start Variance = Calculated Start - Baseline Start{NL}{NL}Lecture :{NL}- 0 = conforme baseline{NL}- > 0 = démarrage plus tardif{NL}- < 0 = démarrage plus tôt", _
        "Purpose: variance between calculated start and baseline start.{NL}Calculation / logic:{NL}- Start Variance = Calculated Start - Baseline Start{NL}{NL}Reading:{NL}- 0 = aligned with baseline{NL}- > 0 = later start{NL}- < 0 = earlier start"
    WBS_Onboarding_AddHelpComment comments, VTS_COL_FINISH_VARIANCE, _
        "Rôle : écart entre la fin calculée et la fin baseline.{NL}Calcul / logique :{NL}- Finish Variance = Calculated Finish - Baseline Finish{NL}{NL}Lecture :{NL}- 0 = conforme baseline{NL}- > 0 = fin plus tardive{NL}- < 0 = fin plus tôt", _
        "Purpose: variance between calculated finish and baseline finish.{NL}Calculation / logic:{NL}- Finish Variance = Calculated Finish - Baseline Finish{NL}{NL}Reading:{NL}- 0 = aligned with baseline{NL}- > 0 = later finish{NL}- < 0 = earlier finish"
    WBS_Onboarding_AddHelpComment comments, VTS_COL_DURATION_VARIANCE, _
        "Rôle : écart entre la durée calculée et la durée baseline.{NL}Calcul / logique :{NL}- Duration Variance = Calculated Duration - Baseline Duration{NL}{NL}Lecture :{NL}- 0 = durée inchangée{NL}- > 0 = durée plus longue{NL}- < 0 = durée plus courte", _
        "Purpose: variance between calculated duration and baseline duration.{NL}Calculation / logic:{NL}- Duration Variance = Calculated Duration - Baseline Duration{NL}{NL}Reading:{NL}- 0 = unchanged duration{NL}- > 0 = longer duration{NL}- < 0 = shorter duration"
    WBS_Onboarding_AddHelpComment comments, VTS_COL_DRIVING_LOGIC, _
        "Rôle : source principale ayant piloté le calcul de la tâche.{NL}Valeurs typiques :{NL}- ACTUAL{NL}- FORECAST{NL}- BASELINE{NL}- DEPENDENCY{NL}- SUMMARY{NL}{NL}Utilité :{NL}- Permet de comprendre rapidement pourquoi la date calculée est celle-ci.", _
        "Purpose: main source driving the task calculation.{NL}Typical values:{NL}- ACTUAL{NL}- FORECAST{NL}- BASELINE{NL}- DEPENDENCY{NL}- SUMMARY{NL}{NL}Use:{NL}- Quickly explains why the calculated date is what it is."
    WBS_Onboarding_AddHelpComment comments, VTS_COL_CRITICAL_PATH, _
        "Rôle : indicateur de chemin critique sur le réseau de planning actuel.{NL}Calcul / logique :{NL}- Basé sur le float total courant.{NL}- Une tâche est critique si son Total Float est inférieur ou égal à 0.{NL}{NL}Utilité :{NL}- Aide à identifier les tâches qui pilotent directement la date projet actuelle.", _
        "Purpose: critical path indicator on the current schedule network.{NL}Calculation / logic:{NL}- Based on current total float.{NL}- A task is critical if its Total Float is less than or equal to 0.{NL}{NL}Use:{NL}- Helps identify tasks directly driving the current project finish date."
    WBS_Onboarding_AddHelpComment comments, VTS_COL_CRITICAL_PATH_REX, _
        "Rôle : indicateur de chemin critique sur l'état de référence Baseline.{NL}Calcul / logique :{NL}- Utilise le même graphe que les analytics Current.{NL}- Analyse les dates Baseline posées, sans compresser les gaps non exprimés en lag.{NL}- Une tâche est critique si son Total Float REX est inférieur ou égal à 0.{NL}{NL}Utilité :{NL}- Sert à l'analyse rétrospective ou comparative sur l'état Baseline.", _
        "Purpose: critical path indicator on the Baseline reference state.{NL}Calculation / logic:{NL}- Uses the same graph as Current analytics.{NL}- Analyzes the placed Baseline dates without compressing gaps that are not expressed as lag.{NL}- A task is critical if its Total Float REX is less than or equal to 0.{NL}{NL}Use:{NL}- Used for retrospective or comparative analysis on the Baseline state."
    WBS_Onboarding_AddHelpComment comments, VTS_COL_LONGEST_PATH, _
        "Rôle : indique si la tâche appartient au plus long chemin du réseau de planning actuel.{NL}Calcul / logique :{NL}- Le moteur marque LONGEST les tâches non terminées reliées à la date de fin du réseau courant par des liens directeurs.{NL}{NL}Utilité :{NL}- Identifie la séquence active la plus longue jusqu'à la fin du projet.", _
        "Purpose: indicates whether the task belongs to the longest path in the current schedule network.{NL}Calculation / logic:{NL}- The engine marks unfinished tasks as LONGEST when driving links connect them to the current network finish.{NL}{NL}Use:{NL}- Identifies the longest active sequence leading to project completion."
    WBS_Onboarding_AddHelpComment comments, VTS_COL_LONGEST_PATH_REX, _
        "Rôle : sortie calculée réservée au plus long chemin sur l'état de référence Baseline / REX.{NL}{NL}Règles :{NL}- Ne pas renseigner manuellement.{NL}- Peut rester vide lorsque le workflow REX ne produit pas cet indicateur.", _
        "Purpose: calculated output reserved for the longest path on the Baseline / REX reference state.{NL}{NL}Rules:{NL}- Do not enter a value manually.{NL}- May remain blank when the REX workflow does not produce this indicator."
    WBS_Onboarding_AddHelpComment comments, VTS_COL_TOTAL_FLOAT, _
        "Rôle : marge totale sur le planning actuel.{NL}Calcul / logique :{NL}- Nombre de jours pendant lesquels la tâche peut glisser sans décaler la date de fin projet actuelle.{NL}{NL}Lecture :{NL}- 0 = critique{NL}- > 0 = marge disponible{NL}- < 0 = float négatif, planning incohérent ou contraint", _
        "Purpose: total float on the current schedule.{NL}Calculation / logic:{NL}- Number of days the task can slip without delaying the current project finish date.{NL}{NL}Reading:{NL}- 0 = critical{NL}- > 0 = available margin{NL}- < 0 = negative float, constrained or inconsistent schedule"
    WBS_Onboarding_AddHelpComment comments, VTS_COL_FREE_FLOAT, _
        "Rôle : marge libre sur le planning actuel.{NL}{NL}Calcul / logique :{NL}- Nombre de jours pendant lesquels la tâche peut glisser sans impacter le début au plus tôt de la tâche suivante.{NL}{NL}Cas particulier :{NL}- Si le float est négatif, cela signifie que la tâche ne respecte pas les contraintes du réseau.{NL}- La date de fin actuelle est déjà trop tardive par rapport aux exigences des successeurs (ou la durée est insuffisante).{NL}{NL}Utilité :{NL}- Mesure la marge locale, plus fine que le Total Float.{NL}- Permet d’identifier immédiatement les incohérences ou contraintes impossibles dans le planning.", _
        "Purpose: free float on the current schedule.{NL}{NL}Calculation / logic:{NL}- Number of days the task can slip without affecting the earliest start of the next task.{NL}{NL}Special case:{NL}- A negative float means the task violates network constraints.{NL}- The current finish is already too late relative to successor requirements (or duration is insufficient).{NL}{NL}Use:{NL}- Measures local margin, more granular than Total Float.{NL}- Helps detect inconsistencies or infeasible constraints in the schedule."
    WBS_Onboarding_AddHelpComment comments, VTS_COL_TOTAL_FLOAT_REX, _
        "Rôle : marge totale sur l'état de référence Baseline / REX.{NL}Calcul / logique :{NL}- Equivalent Baseline du Total Float, calculé avec le même graphe que Current sur les dates Baseline posées.{NL}{NL}Utilité :{NL}- Sert aux comparaisons et au retour d'expérience.", _
        "Purpose: total float on the Baseline / REX reference state.{NL}Calculation / logic:{NL}- Baseline equivalent of Total Float, calculated with the same graph as Current on the placed Baseline dates.{NL}{NL}Use:{NL}- Used for comparisons and lessons learned analysis."
    WBS_Onboarding_AddHelpComment comments, VTS_COL_FREE_FLOAT_REX, _
        "Rôle : marge libre sur l'état de référence Baseline / REX.{NL}Calcul / logique :{NL}- Equivalent Baseline du Free Float, calculé avec le même graphe que Current sur les dates Baseline posées.{NL}{NL}Utilité :{NL}- Sert à l'analyse fine des marges sur l'état Baseline.", _
        "Purpose: free float on the Baseline / REX reference state.{NL}Calculation / logic:{NL}- Baseline equivalent of Free Float, calculated with the same graph as Current on the placed Baseline dates.{NL}{NL}Use:{NL}- Used for detailed float analysis on the Baseline state."
    WBS_Onboarding_AddHelpComment comments, VTS_COL_DEADLINE_FLOAT, _
        "Rôle : marge entre la deadline de la tâche et sa date de fin calculée.{NL}Calcul / logique :{NL}- Deadline Float = Deadline - Calculated Finish.{NL}{NL}Lecture :{NL}- 0 = échéance atteinte exactement.{NL}- > 0 = marge disponible avant l'échéance.{NL}- < 0 = échéance dépassée.", _
        "Purpose: margin between the task deadline and its calculated finish date.{NL}Calculation / logic:{NL}- Deadline Float = Deadline - Calculated Finish.{NL}{NL}Reading:{NL}- 0 = deadline met exactly.{NL}- > 0 = available margin before the deadline.{NL}- < 0 = deadline exceeded."

    Set WBS_Onboarding_BuildHelpCommentMap = comments

End Function

'------------------------------------------------------------------------------
' FR: Ajoute une paire de textes localisés au catalogue sans exposer son stockage.
' EN: Adds one localized text pair to the catalog without exposing its storage.
'------------------------------------------------------------------------------
Private Sub WBS_Onboarding_AddHelpComment( _
    ByVal comments As Object, _
    ByVal columnName As String, _
    ByVal frText As String, _
    ByVal enText As String)

    comments(columnName) = Array( _
        WBS_Onboarding_DecodeHelpText(frText), _
        WBS_Onboarding_DecodeHelpText(enText))

End Sub

'------------------------------------------------------------------------------
' FR: Restaure les sauts de ligne d'un texte d'aide encodé dans le source VBA.
' EN: Restores line breaks in help text encoded in the VBA source.
'------------------------------------------------------------------------------
Private Function WBS_Onboarding_DecodeHelpText(ByVal encodedText As String) As String

    WBS_Onboarding_DecodeHelpText = Replace$(encodedText, "{NL}", vbCrLf)

End Function
' FR: Definit le statut fonctionnel de chaque colonne canonique WBS.
' EN: Defines the functional status of every canonical WBS column.
'------------------------------------------------------------------------------
Private Function WBS_Onboarding_BuildStatusMap() As Object

    Dim statuses As Object
    Dim requiredLabel As String
    Dim optionalLabel As String
    Dim calculatedLabel As String

    requiredLabel = WBS_Onboarding_RequiredLabel()
    optionalLabel = WBS_L("Optionnel", "Optional")
    calculatedLabel = WBS_L("Calculé", "Calculated")

    Set statuses = CreateObject("Scripting.Dictionary")
    statuses.CompareMode = vbTextCompare

    statuses.Add VTS_COL_ID, requiredLabel
    statuses.Add VTS_COL_WBS, requiredLabel
    statuses.Add VTS_COL_TASK_NAME, requiredLabel
    statuses.Add VTS_COL_TASK_DESCRIPTION, optionalLabel
    statuses.Add VTS_COL_DISCIPLINE, optionalLabel
    statuses.Add VTS_COL_SUPPLIER, optionalLabel
    statuses.Add VTS_COL_PROJECT, optionalLabel
    statuses.Add VTS_COL_CAL, optionalLabel
    statuses.Add VTS_COL_TASK_TYPE, optionalLabel
    statuses.Add VTS_COL_S, "O"
    statuses.Add VTS_COL_COMMENTS, optionalLabel
    statuses.Add VTS_COL_PREDECESSORS_WBS, requiredLabel & " G*"
    statuses.Add VTS_COL_WEIGHT_PERCENT, requiredLabel & " S"
    statuses.Add VTS_COL_PROGRESS_PERCENT, requiredLabel & " S"
    statuses.Add VTS_COL_BASELINE_START, requiredLabel & " G*"
    statuses.Add VTS_COL_BASELINE_DURATION, requiredLabel & " G*"
    statuses.Add VTS_COL_BASELINE_FINISH, calculatedLabel
    statuses.Add VTS_COL_ACTUAL_START, optionalLabel
    statuses.Add VTS_COL_ACTUAL_FINISH, optionalLabel
    statuses.Add VTS_COL_ACTUAL_DURATION, calculatedLabel
    statuses.Add VTS_COL_FORECAST_START, optionalLabel
    statuses.Add VTS_COL_FORECAST_FINISH, optionalLabel
    statuses.Add VTS_COL_CALCULATED_START, calculatedLabel
    statuses.Add VTS_COL_CALCULATED_FINISH, calculatedLabel
    statuses.Add VTS_COL_CALCULATED_DURATION, calculatedLabel
    statuses.Add VTS_COL_START_VARIANCE, calculatedLabel
    statuses.Add VTS_COL_FINISH_VARIANCE, calculatedLabel
    statuses.Add VTS_COL_DURATION_VARIANCE, calculatedLabel
    statuses.Add VTS_COL_DRIVING_LOGIC, calculatedLabel
    statuses.Add VTS_COL_CRITICAL_PATH, calculatedLabel
    statuses.Add VTS_COL_CRITICAL_PATH_REX, calculatedLabel
    statuses.Add VTS_COL_LONGEST_PATH, calculatedLabel
    statuses.Add VTS_COL_LONGEST_PATH_REX, calculatedLabel
    statuses.Add VTS_COL_TOTAL_FLOAT, calculatedLabel
    statuses.Add VTS_COL_FREE_FLOAT, calculatedLabel
    statuses.Add VTS_COL_TOTAL_FLOAT_REX, calculatedLabel
    statuses.Add VTS_COL_FREE_FLOAT_REX, calculatedLabel
    statuses.Add VTS_COL_DEADLINE_FLOAT, calculatedLabel

    Set WBS_Onboarding_BuildStatusMap = statuses

End Function

'------------------------------------------------------------------------------
' FR: Retourne le libelle Requis dans la langue runtime WBS.
' EN: Returns the Required label in the WBS runtime language.
'------------------------------------------------------------------------------
Private Function WBS_Onboarding_RequiredLabel() As String

    WBS_Onboarding_RequiredLabel = WBS_L("Requis", "Required")

End Function

'------------------------------------------------------------------------------
' FR: Applique le style compact et met uniquement le statut Requis en evidence.
' EN: Applies the compact style and highlights only the Required status.
'------------------------------------------------------------------------------
Private Sub WBS_Onboarding_FormatStatusCell(ByVal cell As Range)

    Dim requiredStart As Long
    Dim requiredLabel As String

    If cell Is Nothing Then Exit Sub
    requiredLabel = WBS_Onboarding_RequiredLabel()

    cell.Font.Color = RGB(0, 0, 0)
    cell.Font.Bold = True

    requiredStart = InStr(1, CStr(cell.Value2), requiredLabel, vbTextCompare)
    If requiredStart > 0 Then
        With cell.Characters(requiredStart, Len(requiredLabel)).Font
            .Color = RGB(192, 0, 0)
            .Bold = True
        End With
    End If

End Sub

'------------------------------------------------------------------------------
' FR: Affiche le guide général WBS dans O1:R2 en respectant la langue active.
' EN: Displays the general WBS guide in O1:R2 using the active language.
'------------------------------------------------------------------------------
Private Sub WBS_Onboarding_WriteQuickStart(ByVal ws As Worksheet)

    Dim noteArea As Range
    Dim noteText As String
    Dim requiredLabel As String
    Dim requiredStart As Long
    Dim textChanged As Boolean

    Set noteArea = ws.Range("O1:R2")
    noteText = WBS_Onboarding_QuickStartText()
    requiredLabel = WBS_Onboarding_RequiredLabel()

    If Not CBool(noteArea.MergeCells) Or _
        ws.Range("O1").MergeArea.Address(False, False) <> "O1:R2" Then
        noteArea.UnMerge
        noteArea.ClearContents
        noteArea.Merge
        textChanged = True
    End If

    If WBS_Onboarding_NormalizeComparisonText(CStr(ws.Range("O1").Value2)) <> _
        WBS_Onboarding_NormalizeComparisonText(noteText) Then
        noteArea.Value = noteText
        textChanged = True
    End If

    If noteArea.HorizontalAlignment <> xlLeft Then _
        noteArea.HorizontalAlignment = xlLeft
    If noteArea.VerticalAlignment <> xlTop Then _
        noteArea.VerticalAlignment = xlTop
    If Not CBool(noteArea.WrapText) Then noteArea.WrapText = True
    If CBool(noteArea.ShrinkToFit) Then noteArea.ShrinkToFit = False

    If textChanged Then
        requiredStart = InStr(1, noteText, requiredLabel, vbTextCompare)
        If requiredStart > 0 Then
            With noteArea.Characters(requiredStart, Len(requiredLabel)).Font
                .Color = RGB(192, 0, 0)
                .Bold = True
            End With
        End If
    End If

End Sub

'------------------------------------------------------------------------------
' FR: Verifie ou cree WBS Main Buttons si necessaire.
' EN: Ensures or creates WBS Main Buttons when needed.
'------------------------------------------------------------------------------
Public Sub Ensure_WBS_Main_Buttons()

    Dim ws As Worksheet
    Dim startLeft As Double
    Dim topPos As Double
    Dim btnWidthGreen As Double
    Dim btnWidthRed As Double
    Dim btnWidthForced As Double
    Dim btnWidthReset As Double
    Dim btnHeight As Double
    Dim gap As Double

    Set ws = ThisWorkbook.Worksheets("WBS")
    WBS_SetLanguage Settings_GetOwnerLanguage("WBS")

    Ensure_WBS_Onboarding_Guide ws, False

    On Error Resume Next
    ws.Shapes("btn_WBS_FullReset").Delete
    On Error GoTo 0

    Ensure_WBS_TaskType_Input_Setup ws
    Ensure_CriticalPathMode_Toggle

    startLeft = ws.Range("A1").Left + 6
    topPos = 8

    btnWidthGreen = 68
    btnWidthRed = 82
    btnWidthForced = 104
    btnWidthReset = 86
    btnHeight = 34
    gap = 8

    CreateOrUpdateWBSFloatingButton _
        ws, _
        "btn_WBS_Planning", _
        WBS_L("Mettre à jour" & vbCrLf & "Planning", "Update" & vbCrLf & "Planning"), _
        "Run_Planning_Update", _
        startLeft, _
        topPos, _
        btnWidthGreen, _
        btnHeight, _
        RGB(112, 173, 71)

    CreateOrUpdateWBSFloatingButton _
        ws, _
        "btn_WBS_Gantt", _
        WBS_L("Mettre à jour" & vbCrLf & "Gantt", "Update" & vbCrLf & "Gantt"), _
        "Run_Gantt_Update", _
        startLeft + btnWidthGreen + gap, _
        topPos, _
        btnWidthGreen, _
        btnHeight, _
        RGB(112, 173, 71)

    CreateOrUpdateWBSFloatingButton _
        ws, _
        "btn_WBS_SCurve", _
        WBS_L("Mettre à jour" & vbCrLf & "S-Curve", "Update" & vbCrLf & "S-Curve"), _
        "Run_SCurve_Update", _
        startLeft + (btnWidthGreen + gap) * 2, _
        topPos, _
        btnWidthGreen, _
        btnHeight, _
        RGB(112, 173, 71)

    CreateOrUpdateWBSFloatingButton _
        ws, _
        "btn_WBS_ForcedPlanning", _
        WBS_L("MàJ forcée" & vbCrLf & "Planning", "Forced" & vbCrLf & "Planning Update"), _
        "Run_Forced_Planning_Update", _
        startLeft + (btnWidthGreen + gap) * 3, _
        topPos, _
        btnWidthForced, _
        btnHeight, _
        RGB(192, 0, 0)

    CreateOrUpdateWBSFloatingButton _
        ws, _
        "btn_WBS_Full", _
        WBS_L("Mise à jour" & vbCrLf & "complète", "Full" & vbCrLf & "Update"), _
        "Run_Full_Update", _
        startLeft + (btnWidthGreen + gap) * 3 + btnWidthForced + gap, _
        topPos, _
        btnWidthRed, _
        btnHeight, _
        RGB(192, 0, 0)

    CreateOrUpdateWBSFloatingButton _
        ws, _
        "btn_WBS_ResetPlanning", _
        WBS_L("Réinitialiser" & vbCrLf & "Planning", "Reset" & vbCrLf & "Planning"), _
        "Reset_Planning", _
        startLeft + (btnWidthGreen + gap) * 3 + btnWidthForced + gap + btnWidthRed + gap, _
        topPos, _
        btnWidthReset, _
        btnHeight, _
        RGB(192, 0, 0)

    WBS_SetCellValueIfDifferent ws.Range("L1"), _
        WBS_L("Fin baseline" & vbCrLf & "projet", "Project Baseline" & vbCrLf & "Finish")
    WBS_SetCellValueIfDifferent ws.Range("M1"), _
        WBS_L("Fin calculée" & vbCrLf & "projet", "Project" & vbCrLf & "Calculated Finish")
    WBS_SetCellValueIfDifferent ws.Range("N1"), _
        WBS_L("Retard projet" & vbCrLf & "(jours)", "Project Delay" & vbCrLf & "(days)")

End Sub

'------------------------------------------------------------------------------
' FR: Actualise WBS Apply Language sans modifier les regles metier qui produisent les donnees.
' EN: Refreshes WBS Apply Language without changing the business rules that produce the data.
'------------------------------------------------------------------------------

Public Sub WBS_ApplyLanguage(Optional ByVal languageCode As String = "")

    Dim ws As Worksheet
    Dim tbl As ListObject

    On Error GoTo ErrHandler

    Set ws = ThisWorkbook.Worksheets("WBS")

    If Trim$(languageCode) <> "" Then
        WBS_SetLanguage languageCode
    Else
        EnsureWBSLanguageInitialized
    End If

    WBS_SetShapeText ws, "btn_WBS_Planning", WBS_L("Mettre à jour" & vbCrLf & "Planning", "Update" & vbCrLf & "Planning")
    WBS_SetShapeText ws, "btn_WBS_Gantt", WBS_L("Mettre à jour" & vbCrLf & "Gantt", "Update" & vbCrLf & "Gantt")
    WBS_SetShapeText ws, "btn_WBS_SCurve", WBS_L("Mettre à jour" & vbCrLf & "S-Curve", "Update" & vbCrLf & "S-Curve")
    WBS_SetShapeText ws, "btn_WBS_ForcedPlanning", WBS_L("MàJ forcée" & vbCrLf & "Planning", "Forced" & vbCrLf & "Planning Update")
    WBS_SetShapeText ws, "btn_WBS_Full", WBS_L("Mise à jour" & vbCrLf & "complète", "Full" & vbCrLf & "Update")
    WBS_SetShapeText ws, "btn_WBS_ResetPlanning", WBS_L("Réinitialiser" & vbCrLf & "Planning", "Reset" & vbCrLf & "Planning")

    WBS_SetCellValueIfDifferent ws.Range("L1"), _
        WBS_L("Fin baseline" & vbCrLf & "projet", "Project Baseline" & vbCrLf & "Finish")
    WBS_SetCellValueIfDifferent ws.Range("M1"), _
        WBS_L("Fin calculée" & vbCrLf & "projet", "Project" & vbCrLf & "Calculated Finish")
    WBS_SetCellValueIfDifferent ws.Range("N1"), _
        WBS_L("Retard projet" & vbCrLf & "(jours)", "Project Delay" & vbCrLf & "(days)")

    Set tbl = ws.ListObjects("tbl_WBS")
    If Not WBS_Onboarding_LocalizedContentIsCurrent(ws, tbl) Then
        WBS_Onboarding_ApplyLocalizedContent ws, tbl
    End If

    Exit Sub

ErrHandler:
    Err.Raise Err.Number, "WBS_ApplyLanguage", Err.Description

End Sub

'------------------------------------------------------------------------------
' FR: Reinitialise les compteurs runtime utilises uniquement par le harnais WBS.
' EN: Resets runtime counters used only by the WBS proof harness.
'------------------------------------------------------------------------------
Public Sub WBS_OnboardingInstrumentation_Reset()

    gWBSOnboardingLocalizedCheckCount = 0
    gWBSOnboardingLocalizedRebuildCount = 0
    gWBSOnboardingStructuralMutationCount = 0
    gWBSOnboardingThreadedReadCount = 0
    gWBSOnboardingThreadedAddCount = 0
    gWBSOnboardingThreadedClearCount = 0

End Sub

'------------------------------------------------------------------------------
' FR: Expose un snapshot des compteurs runtime au harnais non interactif.
' EN: Exposes a runtime counter snapshot to the noninteractive proof harness.
'------------------------------------------------------------------------------
Public Function WBS_OnboardingInstrumentation_Snapshot() As String

    WBS_OnboardingInstrumentation_Snapshot = _
        "checks=" & CStr(gWBSOnboardingLocalizedCheckCount) & _
        ";rebuilds=" & CStr(gWBSOnboardingLocalizedRebuildCount) & _
        ";structural=" & CStr(gWBSOnboardingStructuralMutationCount) & _
        ";threadedReads=" & CStr(gWBSOnboardingThreadedReadCount) & _
        ";threadedAdds=" & CStr(gWBSOnboardingThreadedAddCount) & _
        ";threadedClears=" & CStr(gWBSOnboardingThreadedClearCount) & _
        ";threadedSupport=" & ExcelCompatibility_ThreadedCommentsProbeStatus()

End Function

'------------------------------------------------------------------------------
' FR: Ecrit une valeur de cellule uniquement lorsqu'elle differe.
' EN: Writes a cell value only when it differs.
'------------------------------------------------------------------------------
Private Sub WBS_SetCellValueIfDifferent( _
    ByVal target As Range, _
    ByVal expectedValue As String)

    If target Is Nothing Then Exit Sub
    If CStr(target.Value2) <> expectedValue Then target.Value = expectedValue

End Sub

'------------------------------------------------------------------------------
' FR: Active ou initialise WBS Set Language dans l'etat runtime du composant.
' EN: Activates or initializes WBS Set Language in the component runtime state.
'------------------------------------------------------------------------------

Public Sub WBS_SetLanguage(ByVal languageCode As String)

    Select Case UCase$(Trim$(languageCode))
        Case "FR"
            gWBSLanguage = "FR"
        Case "EN"
            gWBSLanguage = "EN"
        Case Else
            gWBSLanguage = "EN"
    End Select

End Sub

'------------------------------------------------------------------------------
' FR: Retourne la valeur WBS Current Language sans modifier les donnees d'entree.
' EN: Returns the WBS Current Language value without mutating input data.
'------------------------------------------------------------------------------

Public Function WBS_CurrentLanguage() As String

    EnsureWBSLanguageInitialized
    WBS_CurrentLanguage = gWBSLanguage

End Function

'------------------------------------------------------------------------------
' FR: Verifie ou cree WBSLanguage Initialized si necessaire.
' EN: Ensures or creates WBSLanguage Initialized when needed.
'------------------------------------------------------------------------------
Private Sub EnsureWBSLanguageInitialized()

    If UCase$(Trim$(gWBSLanguage)) <> "FR" And UCase$(Trim$(gWBSLanguage)) <> "EN" Then
        gWBSLanguage = "EN"
    End If

End Sub

'------------------------------------------------------------------------------
' FR: Retourne la valeur WBS L sans modifier les donnees d'entree.
' EN: Returns the WBS L value without mutating input data.
'------------------------------------------------------------------------------

Private Function WBS_L(ByVal frText As String, ByVal enText As String) As String

    If WBS_CurrentLanguage() = "FR" Then
        WBS_L = frText
    Else
        WBS_L = enText
    End If

End Function

'------------------------------------------------------------------------------
' FR: Active ou initialise WBS Set Shape Text dans l'etat runtime du composant.
' EN: Activates or initializes WBS Set Shape Text in the component runtime state.
' FR - Effet de bord : cree ou met a jour des shapes Excel.
' EN - Side effect: creates or updates Excel shapes.
'------------------------------------------------------------------------------

Private Sub WBS_SetShapeText( _
    ByVal ws As Worksheet, _
    ByVal shapeName As String, _
    ByVal captionText As String)

    Dim shp As Shape

    If ws Is Nothing Then Exit Sub

    On Error Resume Next
    Set shp = ws.Shapes(shapeName)
    On Error GoTo 0

    If shp Is Nothing Then Exit Sub

    If CStr(shp.TextFrame2.TextRange.Text) <> captionText Then
        shp.TextFrame2.TextRange.Text = captionText
    End If

End Sub
'------------------------------------------------------------------------------
' FR: Cree Or Update WBSFloating Button pour le domaine WBS buttons.
' EN: Creates Or Update WBSFloating Button for the WBS buttons domain.
'------------------------------------------------------------------------------
Private Sub CreateOrUpdateWBSFloatingButton( _
    ByVal ws As Worksheet, _
    ByVal shpName As String, _
    ByVal captionText As String, _
    ByVal macroName As String, _
    ByVal leftPos As Double, _
    ByVal topPos As Double, _
    ByVal btnWidth As Double, _
    ByVal btnHeight As Double, _
    ByVal fillColor As Long)

    Dim shp As Shape

    On Error Resume Next
    Set shp = ws.Shapes(shpName)
    On Error GoTo 0

    If shp Is Nothing Then
        Set shp = ws.Shapes.AddShape( _
            msoShapeRoundedRectangle, _
            leftPos, _
            topPos, _
            btnWidth, _
            btnHeight)
        shp.Name = shpName
    End If

    shp.Left = leftPos
    shp.Top = topPos
    shp.Width = btnWidth
    shp.Height = btnHeight
    shp.OnAction = macroName
    shp.Placement = xlFreeFloating

    With shp.TextFrame2
        .TextRange.Text = captionText
        .TextRange.Font.Size = 10
        .TextRange.Font.Bold = msoTrue
        .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .VerticalAnchor = msoAnchorMiddle
        .TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .MarginLeft = 6
        .MarginRight = 6
        .MarginTop = 2
        .MarginBottom = 2
    End With

    shp.Fill.ForeColor.RGB = fillColor
    shp.Line.ForeColor.RGB = RGB(150, 150, 150)
    shp.Line.Weight = 1

End Sub

'------------------------------------------------------------------------------
' FR: Verifie ou cree WBS Task Type Input Setup si necessaire.
' EN: Ensures or creates WBS Task Type Input Setup when needed.
'------------------------------------------------------------------------------
Private Sub Ensure_WBS_TaskType_Input_Setup(ByVal ws As Worksheet)

    Dim tbl As ListObject
    Dim rng As Range
    Dim cell As Range
    Dim normalizedValue As String
    Dim rowIndex As Long
    Dim oldEvents As Boolean

    If ws Is Nothing Then Exit Sub

    On Error GoTo SafeExit

    On Error Resume Next
    Set tbl = ws.ListObjects("tbl_WBS")
    On Error GoTo SafeExit

    If tbl Is Nothing Then Exit Sub
    If tbl.DataBodyRange Is Nothing Then Exit Sub

    If Not WBS_TableHasSchemaColumn(tbl, VTS_COL_TASK_TYPE) Then
        Err.Raise vbObjectError + 2310, "Ensure_WBS_TaskType_Input_Setup", _
            "Missing required WBS input column: Task Type"
    End If

    Set rng = SchemaListColumn(tbl, VTS_TABLE_WBS, VTS_COL_TASK_TYPE).DataBodyRange

    oldEvents = Application.EnableEvents
    Application.EnableEvents = False

    For Each cell In rng.Cells

        rowIndex = cell.Row - rng.Row + 1

        If Not WBSButtons_RowHasTaskIdentity(tbl, rowIndex) Then
            If Trim$(CStr(cell.value)) <> "" Then cell.ClearContents
        Else
            normalizedValue = Normalize_WBS_TaskType_Value(cell.value)

            If Trim$(CStr(cell.value)) = "" Then
                cell.value = "Task"
            ElseIf CStr(cell.value) <> normalizedValue Then
                cell.value = normalizedValue
            End If
        End If

    Next cell

    With rng.Validation
        .Delete
        .Add Type:=xlValidateList, _
             AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, _
             Formula1:="Task,Milestone,Level of Effort"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = "Task Type"
        .InputMessage = "Choose: Task, Milestone, or Level of Effort."
        .ErrorTitle = "Invalid Task Type"
        .errorMessage = "Allowed values: Task, Milestone, Level of Effort."
        .ShowInput = True
        .ShowError = True
    End With

    rng.NumberFormat = "@"

SafeExit:
    Application.EnableEvents = oldEvents

    If Err.Number <> 0 Then
        WBSButtons_ShowConsoleError _
            "Ensure_WBS_TaskType_Input_Setup", _
            "Erreur dans Ensure_WBS_TaskType_Input_Setup : " & Err.Description, _
            "Error in Ensure_WBS_TaskType_Input_Setup: " & Err.Description
    End If

End Sub

'------------------------------------------------------------------------------
' FR: Normalise WBS Task Type Value dans un format exploitable.
' EN: Normalizes WBS Task Type Value into a usable format.
'------------------------------------------------------------------------------
Private Function Normalize_WBS_TaskType_Value(ByVal rawValue As Variant) As String

    Dim s As String

    s = UCase$(Trim$(CStr(rawValue)))

    Select Case s

        Case "", "TASK", "STANDARD", "NORMAL"
            Normalize_WBS_TaskType_Value = "Task"

        Case "MILESTONE", "MS", "JALON"
            Normalize_WBS_TaskType_Value = "Milestone"

        Case "LEVEL OF EFFORT", "LOE", "LEVEL-OF-EFFORT", "LEVEL_OF_EFFORT"
            Normalize_WBS_TaskType_Value = "Level of Effort"

        Case Else
            Normalize_WBS_TaskType_Value = CStr(rawValue)

    End Select

End Function

'------------------------------------------------------------------------------
' FR: Retourne la reference WBS Buttons Row Has Task IDentity sans modifier les donnees d'entree.
' EN: Returns the WBS Buttons Row Has Task IDentity reference without mutating input data.
'------------------------------------------------------------------------------

Private Function WBSButtons_RowHasTaskIdentity( _
    ByVal tbl As ListObject, _
    ByVal rowIndex As Long) As Boolean

    Dim idVal As String
    Dim wbsVal As String

    On Error GoTo SafeExit

    If tbl Is Nothing Then Exit Function
    If tbl.DataBodyRange Is Nothing Then Exit Function
    If rowIndex < 1 Or rowIndex > tbl.ListRows.Count Then Exit Function

    If WBS_TableHasSchemaColumn(tbl, VTS_COL_ID) Then
        idVal = Trim$(CStr(SchemaListColumn(tbl, VTS_TABLE_WBS, VTS_COL_ID).DataBodyRange.Cells(rowIndex, 1).value))
    End If

    If WBS_TableHasSchemaColumn(tbl, VTS_COL_WBS) Then
        wbsVal = Trim$(CStr(SchemaListColumn(tbl, VTS_TABLE_WBS, VTS_COL_WBS).DataBodyRange.Cells(rowIndex, 1).value))
    End If

    wbsVal = Replace$(wbsVal, ",", ".")
    WBSButtons_RowHasTaskIdentity = (idVal <> "" Or wbsVal <> "")

SafeExit:
End Function
'------------------------------------------------------------------------------
' FR: Retourne la reference WBS Table Has Column sans modifier les donnees d'entree.
' EN: Returns the WBS Table Has Column reference without mutating input data.
'------------------------------------------------------------------------------

Private Function WBS_TableHasSchemaColumn(ByVal tbl As ListObject, ByVal columnKey As String) As Boolean

    WBS_TableHasSchemaColumn = WBS_TableHasPhysicalColumn( _
        tbl, SchemaCurrentColumnTitle(VTS_TABLE_WBS, columnKey))

End Function

Private Function WBS_TableHasPhysicalColumn(ByVal tbl As ListObject, ByVal columnName As String) As Boolean

    Dim col As ListColumn

    On Error Resume Next
    Set col = tbl.ListColumns(columnName)
    On Error GoTo 0

    WBS_TableHasPhysicalColumn = Not col Is Nothing

End Function

'------------------------------------------------------------------------------
' FR: Projette la collection WBS Buttons Show Console Error vers l'interface autorisee par la politique runtime.
' EN: Projects the WBS Buttons Show Console Error collection to the UI allowed by runtime policy.
'------------------------------------------------------------------------------

Private Sub WBSButtons_ShowConsoleError( _
    ByVal procName As String, _
    ByVal frText As String, _
    ByVal enText As String)

    Dim consoleMessages As Collection

    Set consoleMessages = New Collection

    CalcBridge_AddConsoleMessage consoleMessages, _
        "STOP", _
        "FR:" & vbCrLf & _
        frText & vbCrLf & vbCrLf & _
        "EN:" & vbCrLf & _
        enText

    CalcBridge_ShowPlanningConsole consoleMessages

End Sub













