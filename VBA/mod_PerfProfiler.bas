Attribute VB_Name = "mod_PerfProfiler"
Option Explicit

'===============================================================================
' MODULE : mod_PerfProfiler
' DOMAINE / DOMAIN : Performance / Profiler
'
' FR
' Mesure les scopes et compteurs de performance sans modifier les workflows instruments.
' Ne doit pas contourner les contrats publics des autres domaines.
'
' EN
' Measures performance scopes and counters without changing instrumented workflows.
' Must not bypass public contracts owned by other domains.
'
' CONTRATS / CONTRACTS : Profiler_StartSession, Profiler_StopSession, Profiler_SetEnabled, Profiler_IsEnabled, Profiler_ShouldSuppressUserInterface, Profiler_SetUserInterfaceSuppression, Profiler_BeginScope, Profiler_StartScope
' CALLBACKS EXTERNES / EXTERNAL CALLBACKS : Aucun / None
'===============================================================================


#If VBA7 Then
    Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (ByRef counterValue As Currency) As Long
    Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (ByRef frequencyValue As Currency) As Long
#Else
    Private Declare Function QueryPerformanceCounter Lib "kernel32" (ByRef counterValue As Currency) As Long
    Private Declare Function QueryPerformanceFrequency Lib "kernel32" (ByRef frequencyValue As Currency) As Long
#End If

Private Const PROFILE_REPORT_SHEET As String = "PERF_PROFILE"

Private gProfilerEnabled As Boolean
Private gSessionActive As Boolean
Private gSuppressUserInterface As Boolean
Private gFrequency As Double
Private gSessionStart As Double
Private gSessionElapsed As Double
Private gNextToken As Long
Private gStats As Object
Private gActiveSpans As Object
Private gScopeStack As Collection

'------------------------------------------------------------------------------
' FR: Active ou initialise Profiler Start Session dans l'etat runtime du composant.
' EN: Activates or initializes Profiler Start Session in the component runtime state.
'------------------------------------------------------------------------------

Public Sub Profiler_StartSession(Optional ByVal suppressUserInterface As Boolean = True)

    Profiler_Reset
    gSuppressUserInterface = suppressUserInterface
    gProfilerEnabled = True
    gSessionActive = True
    gSessionStart = Profiler_NowSeconds()

End Sub

'------------------------------------------------------------------------------
' FR: Termine Profiler Stop Session et restaure l'etat runtime possede par le composant.
' EN: Ends Profiler Stop Session and restores runtime state owned by the component.
'------------------------------------------------------------------------------

Public Sub Profiler_StopSession(Optional ByVal writeReport As Boolean = True)

    If Not gSessionActive Then Exit Sub

    gSessionElapsed = Profiler_NowSeconds() - gSessionStart
    gSessionActive = False

    If writeReport Then Profiler_WriteReport

    gProfilerEnabled = False
    gSuppressUserInterface = False

End Sub

'------------------------------------------------------------------------------
' FR: Active ou initialise Profiler Set Enabled dans l'etat runtime du composant.
' EN: Activates or initializes Profiler Set Enabled in the component runtime state.
'------------------------------------------------------------------------------

Public Sub Profiler_SetEnabled(ByVal enabledValue As Boolean)

    If enabledValue Then
        If Not gSessionActive Then Profiler_StartSession False
    Else
        If gSessionActive Then Profiler_StopSession False
        gProfilerEnabled = False
        gSuppressUserInterface = False
    End If

End Sub

'------------------------------------------------------------------------------
' FR: Retourne la valeur Profiler Is Enabled sans modifier les donnees d'entree.
' EN: Returns the Profiler Is Enabled value without mutating input data.
'------------------------------------------------------------------------------

Public Function Profiler_IsEnabled() As Boolean
    Profiler_IsEnabled = (gProfilerEnabled And gSessionActive)
End Function

'------------------------------------------------------------------------------
' FR: Retourne la valeur Profiler Should Suppress User Interface sans modifier les donnees d'entree.
' EN: Returns the Profiler Should Suppress User Interface value without mutating input data.
'------------------------------------------------------------------------------

Public Function Profiler_ShouldSuppressUserInterface() As Boolean
    Profiler_ShouldSuppressUserInterface = gSuppressUserInterface
End Function

'------------------------------------------------------------------------------
' FR: Active ou initialise Profiler Set User Interface Suppression dans l'etat runtime du composant.
' EN: Activates or initializes Profiler Set User Interface Suppression in the component runtime state.
'------------------------------------------------------------------------------

Public Sub Profiler_SetUserInterfaceSuppression(ByVal suppressValue As Boolean)
    gSuppressUserInterface = suppressValue
End Sub

'------------------------------------------------------------------------------
' FR: Active ou initialise Profiler Begin Scope dans l'etat runtime du composant.
' EN: Activates or initializes Profiler Begin Scope in the component runtime state.
'------------------------------------------------------------------------------

Public Function Profiler_BeginScope( _
    ByVal scopeName As String, _
    Optional ByVal category As String = "Procedure") As clsPerfScope

    Dim scope As clsPerfScope
    Dim token As Long

    If Not Profiler_IsEnabled() Then Exit Function

    token = Profiler_StartScope(scopeName, category)
    If token = 0 Then Exit Function

    Set scope = New clsPerfScope
    scope.Init token
    Set Profiler_BeginScope = scope

End Function

'------------------------------------------------------------------------------
' FR: Active ou initialise Profiler Start Scope dans l'etat runtime du composant.
' EN: Activates or initializes Profiler Start Scope in the component runtime state.
'------------------------------------------------------------------------------

Public Function Profiler_StartScope( _
    ByVal scopeName As String, _
    Optional ByVal category As String = "Procedure") As Long

    Dim token As Long
    Dim parentToken As Long
    Dim spanData As Variant

    If Not Profiler_IsEnabled() Then Exit Function

    Profiler_EnsureInfrastructure

    gNextToken = gNextToken + 1
    token = gNextToken

    If gScopeStack.Count > 0 Then parentToken = CLng(gScopeStack(gScopeStack.Count))

    spanData = Array( _
        Trim$(scopeName), _
        Trim$(category), _
        Profiler_NowSeconds(), _
        0#, _
        parentToken)

    gActiveSpans.Add CStr(token), spanData
    gScopeStack.Add token
    Profiler_StartScope = token

End Function

'------------------------------------------------------------------------------
' FR: Termine Profiler End Scope et restaure l'etat runtime possede par le composant.
' EN: Ends Profiler End Scope and restores runtime state owned by the component.
'------------------------------------------------------------------------------

Public Sub Profiler_EndScope(ByVal token As Long)

    Dim spanData As Variant
    Dim parentData As Variant
    Dim elapsed As Double
    Dim selfElapsed As Double
    Dim parentToken As Long

    If token = 0 Then Exit Sub
    If gActiveSpans Is Nothing Then Exit Sub
    If Not gActiveSpans.Exists(CStr(token)) Then Exit Sub

    spanData = gActiveSpans(CStr(token))
    elapsed = Profiler_NowSeconds() - CDbl(spanData(2))
    selfElapsed = elapsed - CDbl(spanData(3))
    If selfElapsed < 0# Then selfElapsed = 0#

    Profiler_UpdateStats CStr(spanData(0)), CStr(spanData(1)), elapsed, selfElapsed

    parentToken = CLng(spanData(4))
    If parentToken <> 0 Then
        If gActiveSpans.Exists(CStr(parentToken)) Then
            parentData = gActiveSpans(CStr(parentToken))
            parentData(3) = CDbl(parentData(3)) + elapsed
            gActiveSpans(CStr(parentToken)) = parentData
        End If
    End If

    gActiveSpans.Remove CStr(token)
    Profiler_RemoveScopeFromStack token

End Sub

'------------------------------------------------------------------------------
' FR: Ecrit ou synchronise Profiler Record Operation dans le stockage possede par le domaine.
' EN: Writes or synchronizes Profiler Record Operation in the store owned by the domain.
'------------------------------------------------------------------------------

Public Sub Profiler_RecordCounter( _
    ByVal counterName As String, _
    ByVal counterValue As Long)

    Dim key As String
    Dim statsData As Variant

    If Not Profiler_IsEnabled() Then Exit Sub
    If counterValue < 1 Then Exit Sub

    key = "Counter|" & counterName
    If gStats.Exists(key) Then
        statsData = gStats(key)
    Else
        statsData = Array(counterName, "Counter", CLng(0), 0#, 0#, 0#, 0#)
    End If

    statsData(2) = CLng(statsData(2)) + counterValue
    gStats(key) = statsData

End Sub

'------------------------------------------------------------------------------
' FR: Ecrit ou synchronise Profiler Record Operation dans le stockage possede par le domaine.
' EN: Writes or synchronizes Profiler Record Operation in the store owned by the domain.
'------------------------------------------------------------------------------

Public Sub Profiler_RecordOperation( _
    ByVal operationName As String, _
    Optional ByVal callCount As Long = 1, _
    Optional ByVal elapsedMilliseconds As Double = 0#)

    Dim i As Long
    Dim elapsedSeconds As Double

    If Not Profiler_IsEnabled() Then Exit Sub
    If callCount < 1 Then Exit Sub

    elapsedSeconds = elapsedMilliseconds / 1000#
    For i = 1 To callCount
        Profiler_UpdateStats operationName, "Operation", elapsedSeconds, elapsedSeconds
    Next i

End Sub

'------------------------------------------------------------------------------
' FR: Ecrit ou synchronise Profiler Write Report dans le stockage possede par le domaine.
' EN: Writes or synchronizes Profiler Write Report in the store owned by the domain.
' FR - Effet de bord : ecrit dans une table Excel detenue par le workflow.
' EN - Side effect: writes to an Excel table owned by the workflow.
'------------------------------------------------------------------------------

Public Sub Profiler_WriteReport(Optional ByVal sheetName As String = PROFILE_REPORT_SHEET)

    Dim ws As Worksheet
    Dim key As Variant
    Dim statsData As Variant
    Dim rowNum As Long
    Dim rootSeconds As Double
    Dim lastRow As Long

    Profiler_EnsureInfrastructure

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = sheetName
    End If

    ws.Cells.Clear
    rootSeconds = gSessionElapsed
    If rootSeconds <= 0# And gSessionStart > 0# Then rootSeconds = Profiler_NowSeconds() - gSessionStart

    ws.Range("A1").Value = "Profiling session elapsed (ms)"
    ws.Range("B1").Value = rootSeconds * 1000#
    ws.Range("A2").Value = "Generated"
    ws.Range("B2").Value = Now
    ws.Range("A4:J4").Value = Array( _
        "Name", "Category", "Calls", "Total ms", "Average ms", _
        "Minimum ms", "Maximum ms", "Self ms", "% session", "Key")

    rowNum = 5
    For Each key In gStats.Keys
        statsData = gStats(key)
        ws.Cells(rowNum, 1).Value = statsData(0)
        ws.Cells(rowNum, 2).Value = statsData(1)
        ws.Cells(rowNum, 3).Value = statsData(2)
        ws.Cells(rowNum, 4).Value = CDbl(statsData(3)) * 1000#
        ws.Cells(rowNum, 5).Value = (CDbl(statsData(3)) / CDbl(statsData(2))) * 1000#
        ws.Cells(rowNum, 6).Value = CDbl(statsData(4)) * 1000#
        ws.Cells(rowNum, 7).Value = CDbl(statsData(5)) * 1000#
        ws.Cells(rowNum, 8).Value = CDbl(statsData(6)) * 1000#
        If rootSeconds > 0# Then ws.Cells(rowNum, 9).Value = CDbl(statsData(3)) / rootSeconds
        ws.Cells(rowNum, 10).Value = CStr(key)
        rowNum = rowNum + 1
    Next key

    lastRow = rowNum - 1
    If lastRow >= 5 Then
        With ws.Sort
            .SortFields.Clear
            .SortFields.Add Key:=ws.Range("D5:D" & lastRow), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
            .SetRange ws.Range("A4:J" & lastRow)
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .Apply
        End With
    End If

    ws.Range("D5:H" & Application.Max(5, lastRow)).NumberFormat = "0.000"
    ws.Range("I5:I" & Application.Max(5, lastRow)).NumberFormat = "0.00%"
    ws.Range("B2").NumberFormat = "dd/mm/yyyy hh:mm:ss"
    ws.Columns("A:J").AutoFit
    ws.Visible = xlSheetVisible

End Sub

'------------------------------------------------------------------------------
' FR: Reinitialise Profiler Reset dans le perimetre possede par le composant.
' EN: Resets Profiler Reset within the state owned by the component.
'------------------------------------------------------------------------------

Public Sub Profiler_Reset()

    Set gStats = CreateObject("Scripting.Dictionary")
    Set gActiveSpans = CreateObject("Scripting.Dictionary")
    Set gScopeStack = New Collection
    gNextToken = 0
    gSessionStart = 0#
    gSessionElapsed = 0#
    gSessionActive = False
    gProfilerEnabled = False
    gSuppressUserInterface = False

End Sub

'------------------------------------------------------------------------------
' FR: Cree ou remet en conformite la map Profiler Ensure Infrastructure de maniere idempotente.
' EN: Creates or restores the Profiler Ensure Infrastructure map to a compliant state idempotently.
'------------------------------------------------------------------------------

Private Sub Profiler_EnsureInfrastructure()

    Dim rawFrequency As Currency

    If gStats Is Nothing Then Set gStats = CreateObject("Scripting.Dictionary")
    If gActiveSpans Is Nothing Then Set gActiveSpans = CreateObject("Scripting.Dictionary")
    If gScopeStack Is Nothing Then Set gScopeStack = New Collection

    If gFrequency <= 0# Then
        QueryPerformanceFrequency rawFrequency
        gFrequency = CDbl(rawFrequency)
    End If

End Sub

'------------------------------------------------------------------------------
' FR: Retourne la valeur Profiler Now Seconds sans modifier les donnees d'entree.
' EN: Returns the Profiler Now Seconds value without mutating input data.
'------------------------------------------------------------------------------

Private Function Profiler_NowSeconds() As Double

    Dim rawCounter As Currency

    Profiler_EnsureInfrastructure
    QueryPerformanceCounter rawCounter
    If gFrequency > 0# Then Profiler_NowSeconds = CDbl(rawCounter) / gFrequency

End Function

'------------------------------------------------------------------------------
' FR: Actualise Profiler Update Stats sans modifier les regles metier qui produisent les donnees.
' EN: Refreshes Profiler Update Stats without changing the business rules that produce the data.
'------------------------------------------------------------------------------

Private Sub Profiler_UpdateStats( _
    ByVal scopeName As String, _
    ByVal category As String, _
    ByVal elapsed As Double, _
    ByVal selfElapsed As Double)

    Dim key As String
    Dim statsData As Variant

    key = category & "|" & scopeName

    If gStats.Exists(key) Then
        statsData = gStats(key)
    Else
        statsData = Array(scopeName, category, CLng(0), 0#, elapsed, 0#, 0#)
    End If

    statsData(2) = CLng(statsData(2)) + 1
    statsData(3) = CDbl(statsData(3)) + elapsed
    If CLng(statsData(2)) = 1 Or elapsed < CDbl(statsData(4)) Then statsData(4) = elapsed
    If elapsed > CDbl(statsData(5)) Then statsData(5) = elapsed
    statsData(6) = CDbl(statsData(6)) + selfElapsed
    gStats(key) = statsData

End Sub

'------------------------------------------------------------------------------
' FR: Reinitialise Profiler Remove Scope From Stack dans le perimetre possede par le composant.
' EN: Resets Profiler Remove Scope From Stack within the state owned by the component.
'------------------------------------------------------------------------------

Private Sub Profiler_RemoveScopeFromStack(ByVal token As Long)

    Dim i As Long

    If gScopeStack Is Nothing Then Exit Sub

    For i = gScopeStack.Count To 1 Step -1
        If CLng(gScopeStack(i)) = token Then
            gScopeStack.Remove i
            Exit Sub
        End If
    Next i

End Sub
