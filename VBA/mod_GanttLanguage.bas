Attribute VB_Name = "mod_GanttLanguage"
Option Explicit

'===============================================================================
' MODULE : mod_GanttLanguage
' DOMAINE / DOMAIN : Gantt
'
' FR
' Possede la langue runtime et les traductions UI du domaine.
' Ne modifie aucun calcul ni etat metier.
'
' EN
' Owns domain runtime language and UI translations.
' Changes no calculation or business state.
'
' CONTRATS / CONTRACTS : Gantt_ApplyLanguage, Gantt_SetLanguage, Gantt_CurrentLanguage
' CALLBACKS EXTERNES / EXTERNAL CALLBACKS : Aucun / None
'===============================================================================


' GANTT FR/EN labels and language state.
'=====================================================

Private Const GANTT_SHEET As String = "GANTT"
Private Const TITLE_ROW As Long = 1
Private Const HEADER_ROW_2 As Long = 4
Private Const COL_WBS As Long = 1
Private Const BTN_RESET_NAME As String = "btn_Gantt_Reset"
Private Const BTN_SCENARIO_NAME As String = "btn_Gantt_Scenario"
Private Const BTN_TEST_NAME As String = "btn_Gantt_Test"
Private Const BTN_LOCK_NAME As String = "btn_Gantt_Lock"
Private Const BTN_VIEW_LEFT_NAME As String = "btn_Gantt_View_Left"
Private Const BTN_SCALE_LEFT_NAME As String = "btn_Gantt_Scale_Left"
Private Const BTN_CONSTRAINT_LEFT_NAME As String = "btn_Gantt_Constraint_Left"
Private Const BTN_CP_LEFT_NAME As String = "btn_Gantt_CP_Left"
Private Const BTN_CP_MULTI_LEFT_NAME As String = "lbl_GANTT_CPMode_Toggle"

Private gGanttLanguage As String

'------------------------------------------------------------------------------
' FR: Execute le helper Gantt  Apply Language dans le workflow de rendu GANTT.
' EN: Runs the Gantt  Apply Language helper in the GANTT rendering workflow.
'------------------------------------------------------------------------------
Public Sub Gantt_ApplyLanguage(Optional ByVal languageCode As String = "")

    Dim ws As Worksheet
    Dim oldEvents As Boolean
    Dim oldInternalWrite As Boolean
    Dim stateCaptured As Boolean
    Dim errorDescription As String

    On Error GoTo ErrHandler

    oldEvents = Application.EnableEvents
    oldInternalWrite = GetGanttInternalWrite()
    stateCaptured = True
    Application.EnableEvents = False
    SetGanttInternalWrite True

    Set ws = ThisWorkbook.Worksheets(GANTT_SHEET)

    If Trim$(languageCode) <> "" Then
        Gantt_SetLanguage languageCode
    Else
        EnsureGanttLanguageInitialized
    End If

    Gantt_SetCellTextIfDifferent ws.cells(TITLE_ROW, COL_WBS), Gantt_L("VUE GANTT", "GANTT VIEW")

    Gantt_SetCellTextIfDifferent ws.Range("A" & HEADER_ROW_2), "WBS"
    Gantt_SetCellTextIfDifferent ws.Range("B" & HEADER_ROW_2), Gantt_L("Nom tâche", "Task Name")
    Gantt_SetCellTextIfDifferent ws.Range("C" & HEADER_ROW_2), Gantt_L("Début", "Start")
    Gantt_SetCellTextIfDifferent ws.Range("D" & HEADER_ROW_2), Gantt_L("Fin", "Finish")
    Gantt_SetCellTextIfDifferent ws.Range("E" & HEADER_ROW_2), Gantt_L("Début test", "Test Start")
    Gantt_SetCellTextIfDifferent ws.Range("F" & HEADER_ROW_2), Gantt_L("Fin test", "Test Finish")
    Gantt_SetCellTextIfDifferent ws.Range("G" & HEADER_ROW_2), Gantt_L("Durée", "Duration")
    Gantt_SetCellTextIfDifferent ws.Range("H" & HEADER_ROW_2), "%"
    Gantt_SetCellTextIfDifferent ws.Range("I" & HEADER_ROW_2), Gantt_L("Test %", "Test %")
    Gantt_SetCellTextIfDifferent ws.Range("J" & HEADER_ROW_2), Gantt_L("Logique", "Logic")

    Gantt_SetShapeText ws, BTN_RESET_NAME, Gantt_L("Réinitialiser", "Reset")
    Gantt_SetShapeText ws, BTN_SCENARIO_NAME, Gantt_L("Scénario", "Scenario")
    Gantt_SetShapeText ws, BTN_TEST_NAME, Gantt_L("Test", "Test")
    Gantt_SetShapeText ws, BTN_LOCK_NAME, Gantt_L("Verrouiller", "Lock")

    Gantt_SetShapeText ws, BTN_VIEW_LEFT_NAME, Gantt_L("Détail / Synthèse", "Detail / Summary")
    Gantt_SetShapeText ws, BTN_SCALE_LEFT_NAME, Gantt_L("Jour / Sem. / Mois", "Day / Week / Month")
    Gantt_SetShapeText ws, BTN_CONSTRAINT_LEFT_NAME, Gantt_L("Contrainte", "Constraint")
    Gantt_SetShapeText ws, BTN_CP_LEFT_NAME, Gantt_L("N/A / Chem. Crit. / Le plus long", "None / Critical Path / Longest Path")
    Gantt_SetShapeText ws, BTN_CP_MULTI_LEFT_NAME, Gantt_L("Unique / Multi-projet", "Single / Multiple Project")

    GoTo SafeExit

ErrHandler:
    errorDescription = Err.Description

SafeExit:
    If stateCaptured Then
        SetGanttInternalWrite oldInternalWrite
        Application.EnableEvents = oldEvents
    End If

    If errorDescription <> "" Then
        CalcBridge_ShowSingleConsoleMessage _
            "STOP", _
            "Erreur dans Gantt_ApplyLanguage : " & errorDescription, _
            "Error in Gantt_ApplyLanguage: " & errorDescription
    End If

End Sub
'------------------------------------------------------------------------------
' FR: Execute le helper Gantt  Set Language dans le workflow de rendu GANTT.
' EN: Runs the Gantt  Set Language helper in the GANTT rendering workflow.
'------------------------------------------------------------------------------
Public Sub Gantt_SetLanguage(ByVal languageCode As String)

    Select Case UCase$(Trim$(languageCode))
        Case "FR"
            gGanttLanguage = "FR"
        Case "EN"
            gGanttLanguage = "EN"
        Case Else
            gGanttLanguage = "EN"
    End Select

End Sub

'------------------------------------------------------------------------------
' FR: Retourne une valeur d'etat ou une reference utilisee par le rendu GANTT.
' EN: Returns a state value or reference used by GANTT rendering.
'------------------------------------------------------------------------------
Public Function Gantt_CurrentLanguage() As String

    EnsureGanttLanguageInitialized
    Gantt_CurrentLanguage = gGanttLanguage

End Function

'------------------------------------------------------------------------------
' FR: Verifie et prepare une ressource GANTT requise avant le rendu ou l'interaction.
' EN: Ensures and prepares a GANTT resource required before rendering or interaction.
'------------------------------------------------------------------------------
Private Sub EnsureGanttLanguageInitialized()

    If UCase$(Trim$(gGanttLanguage)) <> "FR" And UCase$(Trim$(gGanttLanguage)) <> "EN" Then
        gGanttLanguage = "EN"
    End If

End Sub

'------------------------------------------------------------------------------
' FR: Execute le helper Gantt  L dans le workflow de rendu GANTT.
' EN: Runs the Gantt  L helper in the GANTT rendering workflow.
'------------------------------------------------------------------------------
Private Function Gantt_L(ByVal frText As String, ByVal enText As String) As String

    If Gantt_CurrentLanguage() = "FR" Then
        Gantt_L = frText
    Else
        Gantt_L = enText
    End If

End Function

'------------------------------------------------------------------------------
' FR: Execute le helper Gantt  Set Shape Text dans le workflow de rendu GANTT.
' EN: Runs the Gantt  Set Shape Text helper in the GANTT rendering workflow.
'------------------------------------------------------------------------------
Private Sub Gantt_SetShapeText( _
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
' FR: Met a jour un libelle de cellule uniquement lorsque sa valeur differe.
' EN: Updates a cell label only when its value differs.
'------------------------------------------------------------------------------
Private Sub Gantt_SetCellTextIfDifferent( _
    ByVal targetCell As Range, _
    ByVal captionText As String)

    If targetCell Is Nothing Then Exit Sub
    If CStr(targetCell.value) <> captionText Then targetCell.value = captionText

End Sub
