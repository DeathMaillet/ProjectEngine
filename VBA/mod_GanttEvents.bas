Attribute VB_Name = "mod_GanttEvents"
Option Explicit

'===============================================================================
' MODULE : mod_GanttEvents
' DOMAINE / DOMAIN : Gantt
'
' FR
' Traite les changements de cellules Gantt et declenche les transactions de simulation appropriees.
' Ne doit pas contourner les contrats publics des autres domaines.
'
' EN
' Handles Gantt cell changes and triggers the appropriate simulation transactions.
' Must not bypass public contracts owned by other domains.
'
' CONTRATS / CONTRACTS : Handle_Gantt_Change
' CALLBACKS EXTERNES / EXTERNAL CALLBACKS : Handle_Gantt_Change
'===============================================================================


'------------------------------------------------------------------------------
' FR: Traite un changement ou evenement pour Gantt Change.
' EN: Handles a change or event for Gantt Change.
'------------------------------------------------------------------------------
Public Sub Handle_Gantt_Change(ByVal ws As Worksheet, ByVal Target As Range)

    Dim cell As Range
    Dim isValidArea As Boolean
    Dim consoleMessages As Collection
    Dim oldEvents As Boolean
    Dim shouldReconcileWatch As Boolean
    Dim errNumber As Long
    Dim dateFormat As String

    Const COL_TEST_START As Long = 5
    Const COL_TEST_FINISH As Long = 6
    Const COL_TEST_PROGRESS As Long = 9
    Const FIRST_TASK_ROW As Long = 5

    oldEvents = Application.EnableEvents
    On Error GoTo SafeExit

    Set consoleMessages = New Collection
    dateFormat = Settings_GetDateNumberFormat()

    If Target Is Nothing Then Exit Sub

    If GetGanttInternalWrite() Then Exit Sub

    If Target.Row < FIRST_TASK_ROW Then Exit Sub

    shouldReconcileWatch = True
    Application.EnableEvents = False

    For Each cell In Target.Cells

        If cell.Row < FIRST_TASK_ROW Then GoTo NextCell

        isValidArea = IsEditableTestCell(ws, cell)

        If Not isValidArea Then

            Application.Undo

            GanttEvents_AddConsoleMessage consoleMessages, "STOP", _
                "Saisie autorisée uniquement dans les colonnes test jaunes des tâches leaf.", _
                "Input is only allowed in yellow test columns for leaf tasks."

            CalcBridge_ShowPlanningConsole consoleMessages
            GoTo SafeExit

        End If

        If cell.Column = COL_TEST_START Or cell.Column = COL_TEST_FINISH Then

            If Trim$(CStr(cell.value)) <> "" Then

                If Not IsDate(cell.value) Then

                    Application.Undo

                    GanttEvents_AddConsoleMessage consoleMessages, "STOP", _
                        "Date invalide dans une colonne test.", _
                        "Invalid date in a test column."

                    CalcBridge_ShowPlanningConsole consoleMessages
                    GoTo SafeExit

                Else

                    cell.value = CDate(cell.value)
                    cell.NumberFormat = dateFormat

                End If

            End If

        End If

        If cell.Column = COL_TEST_PROGRESS Then

            If Trim$(CStr(cell.value)) <> "" Then

                If Not IsNumeric(cell.value) Then

                    Application.Undo

                    GanttEvents_AddConsoleMessage consoleMessages, "STOP", _
                        "Valeur invalide dans Test %.", _
                        "Invalid value in Test %."

                    CalcBridge_ShowPlanningConsole consoleMessages
                    GoTo SafeExit

                Else

                    NormalizeGanttTestPercentCell cell

                End If

            End If

        End If

NextCell:
    Next cell

SafeExit:
    errNumber = Err.Number
    Application.EnableEvents = oldEvents

    If errNumber <> 0 Then

        If consoleMessages Is Nothing Then Set consoleMessages = New Collection

        GanttEvents_AddConsoleMessage consoleMessages, "STOP", _
            "Erreur VBA dans Handle_Gantt_Change." & vbCrLf & _
            "-> vérifier le dernier bloc modifié dans mod_GanttEvents.", _
            "VBA error in Handle_Gantt_Change." & vbCrLf & _
            "-> check the last edited block in mod_GanttEvents."

        CalcBridge_ShowPlanningConsole consoleMessages

    End If

    If shouldReconcileWatch Then
        On Error Resume Next
        GanttDrag_ReconcileWatchState
        On Error GoTo 0
    End If

End Sub

'------------------------------------------------------------------------------
' FR: Ajoute la collection Gantt Events Add Console Message a la structure cible fournie par l'appelant.
' EN: Adds the Gantt Events Add Console Message collection to the target structure supplied by the caller.
'------------------------------------------------------------------------------

Private Sub GanttEvents_AddConsoleMessage( _
    ByVal consoleMessages As Collection, _
    ByVal msgType As String, _
    ByVal frText As String, _
    ByVal enText As String)

    If consoleMessages Is Nothing Then Exit Sub

    CalcBridge_AddConsoleMessage consoleMessages, msgType, _
        "FR:" & vbCrLf & _
        frText & vbCrLf & vbCrLf & _
        "EN:" & vbCrLf & _
        enText

End Sub




