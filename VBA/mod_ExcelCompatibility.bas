Attribute VB_Name = "mod_ExcelCompatibility"
Option Explicit

' Central, in-memory probes for Excel API capabilities that vary by Office build.

Private Const EXCEL_COMPAT_UNKNOWN As Long = 0
Private Const EXCEL_COMPAT_SUPPORTED As Long = 1
Private Const EXCEL_COMPAT_UNSUPPORTED As Long = 2

Private gThreadedCommentsProbeState As Long
Private gThreadedCommentsOverrideState As Long
Private gThreadedCommentsProbeErrorNumber As Long
Private gThreadedCommentsProbeErrorDescription As String

Public Function ExcelCompatibility_SupportsThreadedComments() As Boolean
    Dim probeCell As Range
    Dim threadedComment As Object
    Dim errorNumber As Long
    Dim errorDescription As String

    If gThreadedCommentsOverrideState <> EXCEL_COMPAT_UNKNOWN Then
        ExcelCompatibility_SupportsThreadedComments = _
            (gThreadedCommentsOverrideState = EXCEL_COMPAT_SUPPORTED)
        Exit Function
    End If

    If gThreadedCommentsProbeState = EXCEL_COMPAT_UNKNOWN Then
        On Error Resume Next
        Set probeCell = ThisWorkbook.Worksheets(1).Range("A1")
        Set threadedComment = probeCell.CommentThreaded
        errorNumber = Err.Number
        errorDescription = Err.Description
        On Error GoTo 0

        gThreadedCommentsProbeErrorNumber = errorNumber
        gThreadedCommentsProbeErrorDescription = errorDescription
        If errorNumber = 0 Then
            gThreadedCommentsProbeState = EXCEL_COMPAT_SUPPORTED
        Else
            gThreadedCommentsProbeState = EXCEL_COMPAT_UNSUPPORTED
        End If
    End If

    ExcelCompatibility_SupportsThreadedComments = _
        (gThreadedCommentsProbeState = EXCEL_COMPAT_SUPPORTED)
End Function

Public Function ExcelCompatibility_ThreadedCommentsProbeStatus() As String
    Dim stateText As String

    Select Case gThreadedCommentsProbeState
        Case EXCEL_COMPAT_SUPPORTED
            stateText = "SUPPORTED"
        Case EXCEL_COMPAT_UNSUPPORTED
            stateText = "UNSUPPORTED"
        Case Else
            stateText = "UNKNOWN"
    End Select

    If gThreadedCommentsOverrideState = EXCEL_COMPAT_SUPPORTED Then
        stateText = "OVERRIDE_SUPPORTED"
    ElseIf gThreadedCommentsOverrideState = EXCEL_COMPAT_UNSUPPORTED Then
        stateText = "OVERRIDE_UNSUPPORTED"
    End If

    ExcelCompatibility_ThreadedCommentsProbeStatus = _
        stateText & ";error=" & CStr(gThreadedCommentsProbeErrorNumber) & _
        ";description=" & gThreadedCommentsProbeErrorDescription
End Function

Public Sub ExcelCompatibility_TestSetThreadedCommentsSupported(ByVal supported As Boolean)
    If supported Then
        gThreadedCommentsOverrideState = EXCEL_COMPAT_SUPPORTED
    Else
        gThreadedCommentsOverrideState = EXCEL_COMPAT_UNSUPPORTED
    End If
End Sub

Public Sub ExcelCompatibility_TestClearThreadedCommentsOverride()
    gThreadedCommentsOverrideState = EXCEL_COMPAT_UNKNOWN
    gThreadedCommentsProbeState = EXCEL_COMPAT_UNKNOWN
    gThreadedCommentsProbeErrorNumber = 0
    gThreadedCommentsProbeErrorDescription = vbNullString
End Sub
