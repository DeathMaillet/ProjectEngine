Attribute VB_Name = "mod_CoreLeafProfiler"
Option Explicit

'===============================================================================
' MODULE : mod_CoreLeafProfiler
' DOMAIN : Core performance evidence
'
' Aggregates per-leaf timings and structural counters only when explicitly
' enabled by an audit harness. The disabled production path is one Boolean test.
'===============================================================================

#If VBA7 Then
    Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (ByRef value As Currency) As Long
    Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (ByRef value As Currency) As Long
#Else
    Private Declare Function QueryPerformanceCounter Lib "kernel32" (ByRef value As Currency) As Long
    Private Declare Function QueryPerformanceFrequency Lib "kernel32" (ByRef value As Currency) As Long
#End If

Private gEnabled As Boolean
Private gFrequency As Double
Private gPhaseTotals As Object
Private gPhaseCalls As Object
Private gCounters As Object

Public Sub CoreLeafProfile_SetEnabled(ByVal enabledValue As Boolean)
    gEnabled = enabledValue
    If enabledValue Then CoreLeafProfile_Reset
End Sub

Public Function CoreLeafProfile_IsEnabled() As Boolean
    CoreLeafProfile_IsEnabled = gEnabled
End Function

Public Sub CoreLeafProfile_Reset()
    Set gPhaseTotals = CreateObject("Scripting.Dictionary")
    Set gPhaseCalls = CreateObject("Scripting.Dictionary")
    Set gCounters = CreateObject("Scripting.Dictionary")
End Sub

Public Function CoreLeafProfile_Timestamp() As Double
    Dim rawCounter As Currency
    Dim rawFrequency As Currency

    If Not gEnabled Then Exit Function
    If gFrequency <= 0# Then
        QueryPerformanceFrequency rawFrequency
        gFrequency = CDbl(rawFrequency)
    End If
    QueryPerformanceCounter rawCounter
    If gFrequency > 0# Then CoreLeafProfile_Timestamp = CDbl(rawCounter) / gFrequency
End Function

Public Sub CoreLeafProfile_AddPhase(ByVal phaseName As String, ByVal startedAt As Double)
    Dim elapsed As Double

    If Not gEnabled Then Exit Sub
    CoreLeafProfile_Ensure
    elapsed = CoreLeafProfile_Timestamp() - startedAt
    If gPhaseTotals.Exists(phaseName) Then
        gPhaseTotals(phaseName) = CDbl(gPhaseTotals(phaseName)) + elapsed
        gPhaseCalls(phaseName) = CLng(gPhaseCalls(phaseName)) + 1
    Else
        gPhaseTotals(phaseName) = elapsed
        gPhaseCalls(phaseName) = CLng(1)
    End If
End Sub

Public Sub CoreLeafProfile_Count(ByVal counterName As String, Optional ByVal amount As Long = 1)
    If Not gEnabled Then Exit Sub
    If amount = 0 Then Exit Sub
    CoreLeafProfile_Ensure
    If gCounters.Exists(counterName) Then
        gCounters(counterName) = CLng(gCounters(counterName)) + amount
    Else
        gCounters(counterName) = amount
    End If
End Sub

Public Sub CoreLeafProfile_ExportTSV( _
    ByVal outputPath As String, _
    ByVal runLabel As String, _
    ByVal datasetSize As Long)

    Dim fileNumber As Integer
    Dim key As Variant
    Dim totalSeconds As Double
    Dim calls As Long

    CoreLeafProfile_Ensure
    fileNumber = FreeFile
    Open outputPath For Output As #fileNumber
    Print #fileNumber, "RunLabel" & vbTab & "DatasetSize" & vbTab & "Kind" & vbTab & _
        "Name" & vbTab & "Calls" & vbTab & "TotalMs" & vbTab & "AverageMs" & vbTab & "Value"

    For Each key In gPhaseTotals.Keys
        totalSeconds = CDbl(gPhaseTotals(CStr(key)))
        calls = CLng(gPhaseCalls(CStr(key)))
        Print #fileNumber, runLabel & vbTab & CStr(datasetSize) & vbTab & "Phase" & vbTab & _
            CStr(key) & vbTab & CStr(calls) & vbTab & _
            CoreLeafProfile_InvariantNumber(totalSeconds * 1000#) & vbTab & _
            CoreLeafProfile_InvariantNumber((totalSeconds * 1000#) / calls) & vbTab
    Next key

    For Each key In gCounters.Keys
        Print #fileNumber, runLabel & vbTab & CStr(datasetSize) & vbTab & "Counter" & vbTab & _
            CStr(key) & vbTab & "0" & vbTab & "0" & vbTab & "0" & vbTab & CStr(gCounters(CStr(key)))
    Next key

    Close #fileNumber
End Sub

Private Sub CoreLeafProfile_Ensure()
    If gPhaseTotals Is Nothing Then Set gPhaseTotals = CreateObject("Scripting.Dictionary")
    If gPhaseCalls Is Nothing Then Set gPhaseCalls = CreateObject("Scripting.Dictionary")
    If gCounters Is Nothing Then Set gCounters = CreateObject("Scripting.Dictionary")
End Sub

Private Function CoreLeafProfile_InvariantNumber(ByVal value As Double) As String
    CoreLeafProfile_InvariantNumber = Replace$(Format$(value, "0.000000"), ",", ".")
End Function
